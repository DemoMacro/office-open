/**
 * Drawing parser for DOCX documents.
 *
 * Parses w:drawing elements and extracts image, chart, or SmartArt data.
 *
 * @module
 */
import type { ChartSpaceOptions } from "@office-open/core";
import { parseOnOff } from "@office-open/core";
import {
  blipDesc,
  convertEmuToPixels,
  customGeometryDesc,
  effectListDesc,
  fillDesc,
  imageTypeFromPath,
  outlineDesc,
  parseAngle,
  parseColorChoice,
  parseNonVisualDrawingProperties,
  pictureLockingDesc,
  presetGeometryDesc,
} from "@office-open/core";
import { chartSpaceDesc } from "@office-open/core/chart";
import { scene3DDesc, shape3DDesc } from "@office-open/core/drawing";
import type {
  BlackWhiteMode,
  NonVisualContentPartPropertiesOptions,
  SourceRectangleOptions,
} from "@office-open/core/drawing";
import { parseNonVisualContentPartProperties } from "@office-open/core/drawing";
import {
  COLOR_CATEGORIES,
  LAYOUT_CATEGORIES,
  STYLE_CATEGORIES,
  parseColorDefinition,
  parseLayoutDefinition,
  parseStyleDefinition,
} from "@office-open/core/smartart";
import { attr, attrBool, attrNum, findChild, findFirst, textOf } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import type { ChartOptions } from "@parts/paragraph/run/chart-run";
import type { PictureOptions } from "@parts/paragraph/run/picture-run";
import type { SmartArtOptions } from "@parts/paragraph/run/smartart-run";
import type { GroupOptions } from "@parts/paragraph/run/wpg-group-run";
import type { ShapeOptions } from "@parts/paragraph/run/wps-shape-run";
import type {
  ChartMediaData,
  ContentPartMediaData,
  GroupChildMediaData,
  MediaData,
  MediaDataTransformation,
  GroupCommonMediaData,
  GroupMediaData,
  ShapeMediaData,
} from "@shared/media";
import type { ContentPartOptions, NonVisualPropertiesOptions } from "@shared/media/data";

import { parseParagraph } from "../../body";
import type { DocxReadContext } from "../../context";
import type { GraphicFrameLocksOptions, GroupShapeLocksOptions } from "./descriptor";
import type { DocPropertiesOptions } from "./doc-properties/doc-properties";
import type {
  Floating,
  HorizontalPositionOptions,
  Margins,
  VerticalPositionOptions,
} from "./floating";
import type { ChildOffset, ChildExtent } from "./inline/graphic/graphic-data/wpg/wpg-group";
import { parseBodyProperties } from "./inline/graphic/graphic-data/wps/body-properties";
import type { NonVisualShapePropertiesOptions } from "./inline/graphic/graphic-data/wps/non-visual-shape-properties";
import type {
  ShapeStyleOptions,
  ShapeStyleReferenceOptions,
  ShapeCoreOptions,
} from "./inline/graphic/graphic-data/wps/wps-shape";
import { TextWrappingType } from "./text-wrap";
import type { TextWrapping, WrapPolygon } from "./text-wrap";

/** Union type for parsed drawing child wrappers. */
export type DrawingChild =
  | { picture: PictureOptions }
  | { chart: ChartOptions }
  | { smartArt: SmartArtOptions }
  | { wpsShape: ShapeOptions }
  | { wpgGroup: GroupOptions }
  | { contentPart: ContentPartOptions };

/**
 * Parse a w:drawing element and dispatch to the correct parser
 * based on the graphicData URI.
 */
export function parseDrawingRun(el: Element, ctx: DocxReadContext): DrawingChild | undefined {
  // A content part nests inside a wpg group (wpg:contentPart, no a:graphic
  // wrapper), so it must be checked before the graphicData dispatch; the
  // wp: spelling is accepted leniently for third-party files.
  const contentPartEl = findFirst(el, "wpg:contentPart") ?? findFirst(el, "wp:contentPart");
  if (contentPartEl) {
    // Group children keep the full ContentPartMediaData (type-discriminated);
    // the paragraph-level payload is the public ContentPartOptions.
    const { type: _type, ...contentPart } = parseContentPart(contentPartEl);
    return { contentPart };
  }

  const graphicData = findFirst(el, "a:graphicData");
  if (!graphicData) return undefined;

  const uri = attr(graphicData, "uri") ?? "";

  if (uri.includes("/chart")) {
    return parseChartDrawing(el, ctx);
  }
  if (uri.includes("/diagram")) {
    return parseSmartArtDrawing(el, ctx);
  }
  if (uri.includes("wordprocessingGroup")) {
    return parseWpgGroupDrawing(el, ctx);
  }
  if (uri.includes("wordprocessingShape")) {
    return parseWpsShapeDrawing(el, ctx);
  }
  return parsePictureRun(el, ctx);
}

/**
 * Extent (EMU→pixels), alt text, and floating properties extracted from a
 * w:drawing's wp:inline or wp:anchor wrapper. Shared by image, wps shape,
 * and wpg group parsing.
 */
interface AnchorInfo {
  width?: number;
  height?: number;
  floating?: Floating;
  altText?: DocPropertiesOptions;
  graphicFrameLocks?: GraphicFrameLocksOptions | null;
  /** wp:effectExtent in raw EMUs — round-tripped verbatim. */
  effectExtent?: { l: number; t: number; r: number; b: number };
}

/** Read wp:cNvGraphicFramePr locking flags. An empty element (no graphicFrameLocks
 *  child) returns `{}` so it round-trips as `<wp:cNvGraphicFramePr/>`. */
function readGraphicFrameLocks(el: Element): GraphicFrameLocksOptions {
  const locks = findChild(el, "a:graphicFrameLocks");
  const result: GraphicFrameLocksOptions = {};
  if (!locks) return result;
  const a = locks.attributes ?? {};
  if (a["noGrp"] !== undefined) result.noGrp = parseOnOff(a["noGrp"]) ?? true;
  if (a["noDrilldown"] !== undefined) result.noDrilldown = parseOnOff(a["noDrilldown"]) ?? true;
  if (a["noSelect"] !== undefined) result.noSelect = parseOnOff(a["noSelect"]) ?? true;
  if (a["noChangeAspect"] !== undefined)
    result.noChangeAspect = parseOnOff(a["noChangeAspect"]) ?? true;
  if (a["noMove"] !== undefined) result.noMove = parseOnOff(a["noMove"]) ?? true;
  if (a["noResize"] !== undefined) result.noResize = parseOnOff(a["noResize"]) ?? true;
  return result as GraphicFrameLocksOptions;
}

/**
 * Read wpg:cNvGrpSpPr/a:grpSpLocks (CT_GroupLocking) into GroupShapeLocksOptions.
 * Returns `undefined` when the group has no locks (Word's default → empty cNvGrpSpPr).
 */
function readGrpSpLocks(cNvGrpSpPr: Element | undefined): GroupShapeLocksOptions | undefined {
  if (!cNvGrpSpPr) return undefined;
  const locks = findChild(cNvGrpSpPr, "a:grpSpLocks");
  if (!locks) return undefined;
  const result: GroupShapeLocksOptions = {};
  const a = locks.attributes ?? {};
  if (a["noGrp"] !== undefined) result.noGrp = parseOnOff(a["noGrp"]) ?? true;
  if (a["noUngrp"] !== undefined) result.noUngrp = parseOnOff(a["noUngrp"]) ?? true;
  if (a["noSelect"] !== undefined) result.noSelect = parseOnOff(a["noSelect"]) ?? true;
  if (a["noRot"] !== undefined) result.noRot = parseOnOff(a["noRot"]) ?? true;
  if (a["noChangeAspect"] !== undefined)
    result.noChangeAspect = parseOnOff(a["noChangeAspect"]) ?? true;
  if (a["noMove"] !== undefined) result.noMove = parseOnOff(a["noMove"]) ?? true;
  if (a["noResize"] !== undefined) result.noResize = parseOnOff(a["noResize"]) ?? true;
  return Object.keys(result).length === 0 ? undefined : (result as GroupShapeLocksOptions);
}

/**
 * Extract {@link AnchorInfo} from the drawing's wp:inline or wp:anchor.
 * Returns `null` when the drawing has neither wrapper.
 */
function parseAnchorOrInline(el: Element): AnchorInfo | null {
  const inline = findFirst(el, "wp:inline");
  const anchor = inline ? undefined : findFirst(el, "wp:anchor");
  const parent = inline ?? anchor;
  if (!parent) return null;

  const info: AnchorInfo = {};

  // Extent (EMU)
  const extent = findChild(parent, "wp:extent");
  if (extent) {
    const cxEmu = attrNum(extent, "cx");
    const cyEmu = attrNum(extent, "cy");
    if (cxEmu !== undefined) info.width = cxEmu;
    if (cyEmu !== undefined) info.height = cyEmu;
  }

  // Effect extent (raw EMUs — round-tripped verbatim, never converted to pixels)
  const ee = findChild(parent, "wp:effectExtent");
  if (ee) {
    info.effectExtent = {
      l: attrNum(ee, "l") ?? 0,
      t: attrNum(ee, "t") ?? 0,
      r: attrNum(ee, "r") ?? 0,
      b: attrNum(ee, "b") ?? 0,
    };
  }

  // Alt text (wp:docPr) — keep the id too so it round-trips verbatim
  const docPr = findChild(parent, "wp:docPr");
  if (docPr) {
    const cNvPrOpts = parseNonVisualDrawingProperties(docPr);
    const id = attr(docPr, "id");
    if (id !== undefined || Object.keys(cNvPrOpts).length > 0) {
      const alt: Partial<DocPropertiesOptions> = { ...cNvPrOpts };
      if (id !== undefined) alt.id = id;
      info.altText = alt as DocPropertiesOptions;
    }
  }

  // Graphic frame locks (wp:cNvGraphicFramePr) — preserved verbatim; null when
  // the source wrapper carries none (so stringify omits the element).
  const cNvGraphicFramePr = findChild(parent, "wp:cNvGraphicFramePr");
  info.graphicFrameLocks = cNvGraphicFramePr ? readGraphicFrameLocks(cNvGraphicFramePr) : null;

  // Floating (anchor only)
  if (anchor && !inline) {
    const floating: Partial<Floating> = {};

    // Margins (distT/distB/distL/distR on wp:anchor)
    const margins: Margins = {};
    const distT = attrNum(anchor, "distT");
    if (distT !== undefined) margins.top = distT;
    const distB = attrNum(anchor, "distB");
    if (distB !== undefined) margins.bottom = distB;
    const distL = attrNum(anchor, "distL");
    if (distL !== undefined) margins.left = distL;
    const distR = attrNum(anchor, "distR");
    if (distR !== undefined) margins.right = distR;
    if (Object.keys(margins).length > 0) floating.margins = margins;

    // Position H/V (relativeFrom + align/posOffset)
    const posH = findChild(anchor, "wp:positionH");
    if (posH) {
      const hp = readPosition(posH);
      if (hp) floating.horizontalPosition = hp as HorizontalPositionOptions;
    }
    const posV = findChild(anchor, "wp:positionV");
    if (posV) {
      const vp = readPosition(posV);
      if (vp) floating.verticalPosition = vp as VerticalPositionOptions;
    }

    // Wrap (element name → TextWrappingType number) + optional side
    const wrap = readWrap(anchor);
    if (wrap) floating.wrap = wrap;

    // Anchor-level flags (stringifyAnchor writes all of these)
    const allowOverlap = attrBool(anchor, "allowOverlap");
    if (allowOverlap !== undefined) floating.allowOverlap = allowOverlap;
    const behindDoc = attrBool(anchor, "behindDoc");
    if (behindDoc !== undefined) floating.behindDocument = behindDoc;
    const locked = attrBool(anchor, "locked");
    if (locked !== undefined) floating.lockAnchor = locked;
    const layoutInCell = attrBool(anchor, "layoutInCell");
    if (layoutInCell !== undefined) floating.layoutInCell = layoutInCell;
    const relativeHeight = attrNum(anchor, "relativeHeight");
    if (relativeHeight !== undefined) floating.zIndex = relativeHeight;

    if (Object.keys(floating).length > 0) info.floating = floating as Floating;
  }

  return info;
}

/**
 * Parse a w:drawing element and return picture data wrapped in { picture: ... }.
 */
export function parsePictureRun(
  el: Element,
  ctx: DocxReadContext,
): { picture: PictureOptions } | undefined {
  const info = parseAnchorOrInline(el);
  if (!info) return undefined;

  // Get graphic → graphicData → blip
  const blip = findFirst(el, "a:blip");
  if (!blip) return undefined;

  const rEmbed = attr(blip, "r:embed");
  if (!rEmbed) return undefined;

  // Resolve the media path against the current part's relationships
  const mediaPath = ctx.resolveRelationship(rEmbed);
  if (!mediaPath) return undefined;

  // Read image data from ZIP
  const imageData = ctx.docx.doc.getRaw(mediaPath);
  if (!imageData) return undefined;

  const type = imageTypeFromPath(mediaPath);

  const imageOpts: Record<string, unknown> = {
    type,
    data: imageData,
    // Pin the source file name: type normalization (jpeg→jpg) would otherwise
    // rewrite the extension and drop the source [Content_Types] Default entry.
    fileName: mediaPath.split("/").pop() ?? mediaPath,
    transformation: {
      ...(info.width !== undefined ? { width: info.width } : {}),
      ...(info.height !== undefined ? { height: info.height } : {}),
      ...(info.effectExtent ? { effectExtent: info.effectExtent } : {}),
    },
  };
  if (info.altText) imageOpts.altText = info.altText;
  if (info.floating) imageOpts.floating = info.floating;
  if (info.graphicFrameLocks !== undefined) imageOpts.graphicFrameLocks = info.graphicFrameLocks;

  // Blip-fill crop (pic:blipFill/a:srcRect)
  const blipFill = findFirst(el, "pic:blipFill");
  if (blipFill) {
    const srcRect = readSourceRectangle(blipFill);
    if (srcRect) imageOpts.sourceRectangle = srcRect;
  }

  // Picture non-visual properties (pic:nvPicPr/pic:cNvPr)
  const cNvPr = readPicCnvPr(el, ctx);
  if (cNvPr) imageOpts.nonVisualProperties = cNvPr;

  // Picture shape properties (pic:spPr): outline + fill + effects round-trip
  // via the shared core descriptors (bidirectional).
  const picSpPr = findFirst(el, "pic:spPr");
  if (picSpPr) {
    const fill = readShapeFill(picSpPr, ctx);
    if (fill) imageOpts.fill = fill;
    const ln = findChild(picSpPr, "a:ln");
    if (ln) imageOpts.outline = outlineDesc.parse(ln, ctx);
    const effectLst = findChild(picSpPr, "a:effectLst");
    if (effectLst) imageOpts.effects = effectListDesc.parse(effectLst, ctx);
    // Rotation/flip live on pic:spPr/a:xfrm (ST_Angle in 1/60000 deg). Convert
    // to degrees to match the MediaTransformation API — createTransformation
    // multiplies back by 60_000 on stringify, so integer-degree rotation stays
    // lossless across round-trip.
    const xfrm = findChild(picSpPr, "a:xfrm");
    if (xfrm) {
      const transform = imageOpts.transformation as {
        rotation?: number;
        flip?: { horizontal?: boolean; vertical?: boolean };
      };
      const rot = attrNum(xfrm, "rot");
      if (rot !== undefined) transform.rotation = parseAngle(rot);
      const flipH = attrBool(xfrm, "flipH");
      const flipV = attrBool(xfrm, "flipV");
      if (flipH !== undefined || flipV !== undefined) {
        transform.flip = {
          ...(flipH !== undefined ? { horizontal: flipH } : {}),
          ...(flipV !== undefined ? { vertical: flipV } : {}),
        };
      }
    }
  }

  // Blip recolor effects (a:lum/a:hsl/a:tint/...) under a:blip — image
  // brightness/contrast/tint adjustments applied directly to the image data.
  const blipResult = blipDesc.parse(blip, ctx);
  if (blipResult.blipEffects) imageOpts.blipEffects = blipResult.blipEffects;

  // Blip extension: a14:useLocalDpi (rendering hint, round-trip verbatim).
  const useLocalDpi = readBlipUseLocalDpi(blip);
  if (useLocalDpi !== undefined) imageOpts.useLocalDpi = useLocalDpi;

  // Blip extension: asvg:svgBlip — when present, the a:blip r:embed is the
  // raster fallback and the SVG part is referenced here. Restructure into an
  // SvgMediaOptions (vector primary + raster fallback) so stringify re-emits
  // both branches; otherwise the SVG is dropped on round-trip.
  const svg = readBlipSvg(blip, ctx);
  if (svg) {
    imageOpts.fallback = { type, data: imageData, fileName: imageOpts.fileName };
    imageOpts.type = "svg";
    imageOpts.data = svg.data;
    // The vector part keeps the svg source name; the raster fallback keeps the
    // blip's original media name.
    imageOpts.fileName = svg.fileName;
  }

  return { picture: imageOpts as unknown as PictureOptions };
}

/**
 * Read the `a14:useLocalDpi` blip extension (val="0" → false, "1" → true).
 * Returns undefined when the blip has no useLocalDpi extension.
 */
function readBlipUseLocalDpi(blip: Element): boolean | undefined {
  const extLst = findChild(blip, "a:extLst");
  if (!extLst) return undefined;
  for (const ext of extLst.elements ?? []) {
    if (ext.type !== "element" || ext.name !== "a:ext") continue;
    const useLocalDpiEl = findChild(ext, "a14:useLocalDpi");
    if (useLocalDpiEl) {
      const val = useLocalDpiEl.attributes?.["val"];
      return parseOnOff(val) ?? true;
    }
  }
  return undefined;
}

/**
 * Read the `asvg:svgBlip` blip extension. When present, the surrounding
 * `a:blip` r:embed carries the raster fallback and this extension targets the
 * vector SVG part. Returns the SVG bytes so the picture round-trips as an
 * SvgMediaOptions (vector primary + raster fallback); undefined when no SVG
 * extension exists.
 */
function readBlipSvg(
  blip: Element,
  ctx: DocxReadContext,
): { data: Uint8Array; fileName: string } | undefined {
  const extLst = findChild(blip, "a:extLst");
  if (!extLst) return undefined;
  for (const ext of extLst.elements ?? []) {
    if (ext.type !== "element" || ext.name !== "a:ext") continue;
    const svgBlip = findChild(ext, "asvg:svgBlip");
    if (svgBlip) {
      const rEmbed = attr(svgBlip, "r:embed");
      if (!rEmbed) return undefined;
      const svgPath = ctx.resolveRelationship(rEmbed);
      if (!svgPath) return undefined;
      const data = ctx.docx.doc.getRaw(svgPath);
      if (!data) return undefined;
      return { data, fileName: svgPath.split("/").pop() ?? svgPath };
    }
  }
  return undefined;
}

// ── WPS shape / WPG group parsing ───────────────────────────────────────────

/**
 * Read the blip-fill crop rectangle (`a:srcRect`, l/t/r/b percentage insets)
 * from a `pic:blipFill` parent. Returns undefined when there is no crop.
 */
function readSourceRectangle(parent: Element): SourceRectangleOptions | undefined {
  const sr = findChild(parent, "a:srcRect");
  if (!sr) return undefined;
  const result: SourceRectangleOptions = {};
  const left = attrNum(sr, "l");
  const top = attrNum(sr, "t");
  const right = attrNum(sr, "r");
  const bottom = attrNum(sr, "b");
  if (left !== undefined) result.left = left;
  if (top !== undefined) result.top = top;
  if (right !== undefined) result.right = right;
  if (bottom !== undefined) result.bottom = bottom;
  // An empty <a:srcRect/> is meaningful (explicit no-crop reset), so return
  // the object even when no l/t/r/b attributes are present.
  return result as SourceRectangleOptions;
}

/**
 * Read pic:cNvPr (id/name/descr) from a drawing's pic:nvPicPr. Returns
 * undefined when there is no non-visual properties block.
 */
function readPicCnvPr(el: Element, ctx: DocxReadContext): NonVisualPropertiesOptions | undefined {
  const nvPicPr = findFirst(el, "pic:nvPicPr");
  if (!nvPicPr) return undefined;
  const cNvPr = findChild(nvPicPr, "pic:cNvPr");
  const result: NonVisualPropertiesOptions = { ...parseNonVisualDrawingProperties(cNvPr) };
  if (cNvPr) {
    const id = attrNum(cNvPr, "id");
    if (id !== undefined) result.id = id;
  }
  // pic:cNvPicPr sibling — preferRelativeResize (Word omits the default true;
  // an explicit false round-trips as "0") and the a:picLocks tri-state: null
  // when the source carried none so stringify keeps the bare element.
  const cNvPicPr = findChild(nvPicPr, "pic:cNvPicPr");
  if (cNvPicPr) {
    const preferRelativeResize = attrBool(cNvPicPr, "preferRelativeResize");
    if (preferRelativeResize !== undefined) result.preferRelativeResize = preferRelativeResize;
    const locksEl = findChild(cNvPicPr, "a:picLocks");
    result.pictureLocks = locksEl ? pictureLockingDesc.parse(locksEl, ctx) : null;
  }
  return Object.keys(result).length > 0 ? result : undefined;
}

/**
 * Read a fill element from a shape-properties parent, if present. Delegates to
 * core {@link fillDesc} so solid/gradient/pattern/group/no fills all round-trip
 * through the shared descriptor.
 */
function readShapeFill(parent: Element, ctx: DocxReadContext) {
  const fillChild =
    findChild(parent, "a:noFill") ??
    findChild(parent, "a:solidFill") ??
    findChild(parent, "a:gradFill") ??
    findChild(parent, "a:pattFill") ??
    findChild(parent, "a:grpFill") ??
    findChild(parent, "a:blipFill");
  if (!fillChild) return undefined;
  return fillDesc.parse(parent, ctx);
}

/**
 * Read a single style-matrix reference (a:lnRef/a:fillRef/a:effectRef/a:fontRef):
 * the `idx` attribute plus an optional EG_ColorChoice color override.
 */
function parseStyleRef(el: Element, ctx: DocxReadContext): ShapeStyleReferenceOptions | undefined {
  const idx = attr(el, "idx");
  if (idx === undefined) return undefined;
  const result: ShapeStyleReferenceOptions = { idx };
  const color = parseColorChoice(el, ctx);
  if (color && Object.keys(color).length > 0) result.color = color;
  return result as ShapeStyleReferenceOptions;
}

/**
 * Parse a wps:style (CT_ShapeStyle): line/fill/effect/font references into the
 * document theme. Delegates color to the shared core {@link parseColorChoice}.
 */
function parseShapeStyle(styleEl: Element, ctx: DocxReadContext): ShapeStyleOptions {
  const result: ShapeStyleOptions = {};
  const lnRef = findChild(styleEl, "a:lnRef");
  if (lnRef) result.lineReference = parseStyleRef(lnRef, ctx);
  const fillRef = findChild(styleEl, "a:fillRef");
  if (fillRef) result.fillReference = parseStyleRef(fillRef, ctx);
  const effectRef = findChild(styleEl, "a:effectRef");
  if (effectRef) result.effectReference = parseStyleRef(effectRef, ctx);
  const fontRef = findChild(styleEl, "a:fontRef");
  if (fontRef) result.fontReference = parseStyleRef(fontRef, ctx);
  return result as ShapeStyleOptions;
}

/**
 * Parse the shared core of a `wps:wsp` element (everything except the outer
 * drawing transformation/floating): text content, body properties, fill, and
 * the txBox non-visual flag. Used both for standalone wps shapes and for wps
 * children nested inside a wpg group.
 */
function parseWpsShapeCore(wspEl: Element, ctx: DocxReadContext): ShapeCoreOptions {
  const result: Partial<ShapeCoreOptions> = {};

  // Text content — w:txbxContent (w namespace, per CT_TxbxContent → w:EG_BlockLevelElts)
  // holds the shape's paragraphs, even when wrapped in wps:txbx.
  const txbxContent = findFirst(wspEl, "w:txbxContent");
  const children: ShapeCoreOptions["children"] = [];
  if (txbxContent) {
    for (const child of txbxContent.elements ?? []) {
      if (child.name === "w:p") children.push(parseParagraph(child, ctx));
    }
  }
  result.children = children;

  // Linked text box chain (wps:linkedTxbx) — XSD choice partner of txbx.
  const linkedTxbx = findChild(wspEl, "wps:linkedTxbx");
  if (linkedTxbx) {
    result.linkedTextBox = {
      id: attrNum(linkedTxbx, "id") ?? 0,
      sequence: attrNum(linkedTxbx, "seq") ?? 1,
    };
  }
  // East-Asian vertical flow (wps:wsp @normalEastAsianFlow)
  const normalEastAsianFlow = attrBool(wspEl, "normalEastAsianFlow");
  if (normalEastAsianFlow !== undefined) result.normalEastAsianFlow = normalEastAsianFlow;

  // Non-visual shape properties: wps:cNvPr (id/name/descr) + a choice of
  // wps:cNvSpPr (txBox marker) or wps:cNvCnPr (connector) — mutually exclusive.
  const cNvPr = findChild(wspEl, "wps:cNvPr");
  const cNvSpPr = findChild(wspEl, "wps:cNvSpPr");
  const cNvCnPr = findChild(wspEl, "wps:cNvCnPr");
  const txBox = cNvSpPr ? attr(cNvSpPr, "txBox") : undefined;
  if (cNvPr || txBox !== undefined || cNvCnPr) {
    const nvp: NonVisualShapePropertiesOptions = {};
    if (cNvPr) {
      const id = attrNum(cNvPr, "id");
      const name = attr(cNvPr, "name");
      const descr = attr(cNvPr, "descr");
      const title = attr(cNvPr, "title");
      if (id !== undefined) nvp.id = id;
      if (name) nvp.name = name;
      if (descr) nvp.description = descr;
      if (title) nvp.title = title;
    }
    if (cNvCnPr) nvp.connector = true;
    else if (txBox !== undefined) nvp.textBox = txBox;
    result.nonVisualProperties = nvp as NonVisualShapePropertiesOptions;
  }

  // Shape properties (wps:spPr) — fill/outline/effects/geometry round-trip via
  // the shared core descriptors (bidirectional) so spPr stays structured.
  const spPr = findChild(wspEl, "wps:spPr");
  if (spPr) {
    const fill = readShapeFill(spPr, ctx);
    if (fill) result.fill = fill;
    const ln = findChild(spPr, "a:ln");
    if (ln) result.outline = outlineDesc.parse(ln, ctx);
    const effectLst = findChild(spPr, "a:effectLst");
    if (effectLst) result.effects = effectListDesc.parse(effectLst, ctx);
    const custGeom = findChild(spPr, "a:custGeom");
    if (custGeom) result.customGeometry = customGeometryDesc.parse(custGeom, ctx);
    const prstGeom = findChild(spPr, "a:prstGeom");
    if (prstGeom) result.presetGeometry = presetGeometryDesc.parse(prstGeom, ctx);
    const scene3d = findChild(spPr, "a:scene3d");
    if (scene3d) result.scene3d = scene3DDesc.parse(scene3d, ctx);
    const sp3d = findChild(spPr, "a:sp3d");
    if (sp3d) result.shape3d = shape3DDesc.parse(sp3d, ctx);
  }

  // Body properties (wps:bodyPr)
  const bodyPr = findChild(wspEl, "wps:bodyPr");
  if (bodyPr) result.bodyProperties = parseBodyProperties(bodyPr, ctx);

  // Shape style (wps:style) — theme references (lnRef/fillRef/effectRef/fontRef)
  const styleEl = findChild(wspEl, "wps:style");
  if (styleEl) result.style = parseShapeStyle(styleEl, ctx);

  return result as ShapeCoreOptions;
}

/**
 * Build a child's MediaDataTransformation directly from an `a:xfrm`, keeping
 * EMU values intact (no pixel rounding) so group child coordinates survive
 * round-trip without drift.
 */
function readChildTransformation(spPr: Element | undefined): MediaDataTransformation {
  const result: MediaDataTransformation = {
    pixels: { x: 0, y: 0 },
    emus: { x: 0, y: 0 },
  };
  if (!spPr) return result;
  const xfrm = findChild(spPr, "a:xfrm");
  if (!xfrm) return result;

  const off = findChild(xfrm, "a:off");
  if (off?.attributes) {
    const x = Number(off.attributes["x"] ?? 0);
    const y = Number(off.attributes["y"] ?? 0);
    result.offset = {
      emus: { x, y },
      pixels: { x: convertEmuToPixels(x), y: convertEmuToPixels(y) },
    };
  }
  const ext = findChild(xfrm, "a:ext");
  if (ext?.attributes) {
    const cx = Number(ext.attributes["cx"] ?? 0);
    const cy = Number(ext.attributes["cy"] ?? 0);
    result.emus = { x: cx, y: cy };
    result.pixels = { x: convertEmuToPixels(cx), y: convertEmuToPixels(cy) };
  }

  const flipH = attrBool(xfrm, "flipH");
  const flipV = attrBool(xfrm, "flipV");
  if (flipH !== undefined || flipV !== undefined) {
    const flip: { horizontal?: boolean; vertical?: boolean } = {};
    if (flipH !== undefined) flip.horizontal = flipH;
    if (flipV !== undefined) flip.vertical = flipV;
    result.flip = flip;
  }
  const rot = attrNum(xfrm, "rot");
  if (rot !== undefined) result.rotation = rot;

  return result;
}

/**
 * Build a MediaDataTransformation from a directly nested xfrm (graphicFrame /
 * contentPart carry wp:/wpg:xfrm as a direct child instead of inside spPr).
 */
function readDirectXfrmTransformation(el: Element): MediaDataTransformation {
  const xfrm = (el.elements ?? []).find(
    (c) => c.type === "element" && /^(wp|wpg|a):xfrm$/.test(c.name ?? ""),
  ) as Element | undefined;
  if (!xfrm) return { pixels: { x: 0, y: 0 }, emus: { x: 0, y: 0 } };

  const result: MediaDataTransformation = { pixels: { x: 0, y: 0 }, emus: { x: 0, y: 0 } };
  const off = findChild(xfrm, "a:off");
  if (off?.attributes) {
    result.offset = {
      emus: { x: Number(off.attributes["x"] ?? 0), y: Number(off.attributes["y"] ?? 0) },
      pixels: {
        x: convertEmuToPixels(Number(off.attributes["x"] ?? 0)),
        y: convertEmuToPixels(Number(off.attributes["y"] ?? 0)),
      },
    };
  }
  const ext = findChild(xfrm, "a:ext");
  if (ext?.attributes) {
    const cx = Number(ext.attributes["cx"] ?? 0);
    const cy = Number(ext.attributes["cy"] ?? 0);
    result.emus = { x: cx, y: cy };
    result.pixels = { x: convertEmuToPixels(cx), y: convertEmuToPixels(cy) };
  }
  return result;
}

/**
 * Parse a `wps:wsp` nested inside a wpg group into a {@link ShapeMediaData} child
 * (its transformation kept as EMU via {@link readChildTransformation}).
 */
function parseWpsChildMediaData(wspEl: Element, ctx: DocxReadContext): ShapeMediaData | undefined {
  const data = parseWpsShapeCore(wspEl, ctx);
  const spPr = findChild(wspEl, "wps:spPr");
  return {
    type: "wps",
    transformation: readChildTransformation(spPr),
    data,
  };
}

/**
 * Parse a `pic:pic` nested inside a wpg group into a {@link MediaData} child.
 * Uses the original media path as the registration key so repeated references
 * to the same image collapse to one media entry.
 */
function parsePicChildMediaData(picEl: Element, ctx: DocxReadContext): MediaData | undefined {
  const blip = findFirst(picEl, "a:blip");
  if (!blip) return undefined;
  const rEmbed = attr(blip, "r:embed");
  if (!rEmbed) return undefined;

  const mediaPath = ctx.resolveRelationship(rEmbed);
  if (!mediaPath) return undefined;
  const data = ctx.docx.doc.getRaw(mediaPath);
  if (!data) return undefined;

  const spPr = findChild(picEl, "pic:spPr");
  const result: MediaData = {
    type: imageTypeFromPath(mediaPath),
    // fileName is the bare basename; the compiler writes it under word/media/.
    fileName: mediaPath.split("/").pop() ?? mediaPath,
    data,
    transformation: readChildTransformation(spPr),
  };
  const blipFill = findChild(picEl, "pic:blipFill");
  if (blipFill) {
    const srcRect = readSourceRectangle(blipFill);
    if (srcRect) result.sourceRectangle = srcRect;
  }
  const cNvPr = readPicCnvPr(picEl, ctx);
  if (cNvPr) result.nonVisualProperties = cNvPr;
  // Grouped picture spPr (fill/outline) rides on GroupCommonMediaData so it
  // round-trips through stringifyGroupChild → stringifyShapeProps.
  if (spPr) {
    const fill = readShapeFill(spPr, ctx);
    if (fill) (result as MediaData & GroupCommonMediaData).fill = fill;
    const ln = findChild(spPr, "a:ln");
    if (ln) (result as MediaData & GroupCommonMediaData).outline = outlineDesc.parse(ln, ctx);
  }
  // asvg:svgBlip extension — when present, the a:blip r:embed is the raster
  // fallback and the vector SVG lives in the extension. Reshape into an
  // SvgMediaData so stringify re-emits both (vector + fallback); otherwise the
  // SVG is dropped on round-trip.
  const svg = readBlipSvg(blip, ctx);
  if (svg) {
    return {
      ...result,
      type: "svg",
      data: svg.data,
      fileName: svg.fileName,
      fallback: {
        type: result.type,
        fileName: result.fileName,
        data,
        transformation: result.transformation,
      },
    } as MediaData;
  }
  return result;
}

/**
 * Parse a standalone wps shape drawing (graphicData URI wordprocessingShape).
 */
function parseWpsShapeDrawing(
  el: Element,
  ctx: DocxReadContext,
): { wpsShape: ShapeOptions } | undefined {
  const wsp = findFirst(el, "wps:wsp");
  if (!wsp) return undefined;

  const info = parseAnchorOrInline(el) ?? {};
  const data = parseWpsShapeCore(wsp, ctx);

  const shape: ShapeOptions = {
    ...data,
    transformation: {
      width: info.width ?? 0,
      height: info.height ?? 0,
      ...(info.effectExtent ? { effectExtent: info.effectExtent } : {}),
    },
  };
  if (info.floating) shape.floating = info.floating;
  if (info.altText) shape.altText = info.altText;
  if (info.graphicFrameLocks !== undefined) shape.graphicFrameLocks = info.graphicFrameLocks;

  return { wpsShape: shape as ShapeOptions };
}

/**
 * Parse a wpg group drawing (graphicData URI wordprocessingGroup).
 */
function parseWpgGroupDrawing(
  el: Element,
  ctx: DocxReadContext,
): { wpgGroup: GroupOptions } | undefined {
  const wgp = findFirst(el, "wpg:wgp");
  if (!wgp) return undefined;

  const info = parseAnchorOrInline(el) ?? {};
  const grpSpPr = findChild(wgp, "wpg:grpSpPr");
  const { childOffset, childExtent } = readGroupCoords(grpSpPr);

  const group: GroupOptions = {
    children: parseGroupChildren(wgp, ctx),
    transformation: {
      width: info.width ?? 0,
      height: info.height ?? 0,
      ...(info.effectExtent ? { effectExtent: info.effectExtent } : {}),
    },
  };
  if (childOffset) group.childOffset = childOffset;
  if (childExtent) group.childExtent = childExtent;
  if (info.floating) group.floating = info.floating;
  if (info.altText) group.altText = info.altText;
  if (info.graphicFrameLocks !== undefined) group.graphicFrameLocks = info.graphicFrameLocks;
  const grpSpLocks = readGrpSpLocks(findChild(wgp, "wpg:cNvGrpSpPr"));
  if (grpSpLocks) group.groupShapeLocks = grpSpLocks;
  // Group shape props (grpSpPr): fill + effects round-trip via shared descriptors.
  if (grpSpPr) {
    const fill = readShapeFill(grpSpPr, ctx);
    if (fill) group.fill = fill;
    const effectLst = findChild(grpSpPr, "a:effectLst");
    if (effectLst) group.effects = effectListDesc.parse(effectLst, ctx);
  }

  return { wpgGroup: group as GroupOptions };
}

/**
 * Read chOff/chExt child coordinate space from a group's grpSpPr/a:xfrm.
 * Shared by the top-level wpg:wgp and nested wpg:grpSp.
 */
function readGroupCoords(grpSpPr: Element | undefined): {
  childOffset?: ChildOffset;
  childExtent?: ChildExtent;
} {
  if (!grpSpPr) return {};
  const xfrm = findChild(grpSpPr, "a:xfrm");
  if (!xfrm) return {};
  let childOffset: ChildOffset | undefined;
  let childExtent: ChildExtent | undefined;
  const off = findChild(xfrm, "a:chOff");
  if (off?.attributes) {
    childOffset = { x: Number(off.attributes["x"] ?? 0), y: Number(off.attributes["y"] ?? 0) };
  }
  const ext = findChild(xfrm, "a:chExt");
  if (ext?.attributes) {
    childExtent = { cx: Number(ext.attributes["cx"] ?? 0), cy: Number(ext.attributes["cy"] ?? 0) };
  }
  return { childOffset, childExtent };
}

/**
 * Parse the children of a group element (CT_WordprocessingGroup choice):
 * wps:wsp shapes, pic:pic pictures, and nested wpg:grpSp groups (recursive).
 */
function parseGroupChildren(groupEl: Element, ctx: DocxReadContext): GroupChildMediaData[] {
  const children: GroupChildMediaData[] = [];
  for (const child of groupEl.elements ?? []) {
    if (child.type !== "element") continue;
    const md = parseGroupChild(child, ctx);
    if (md) children.push(md);
  }
  return children;
}

function parseGroupChild(el: Element, ctx: DocxReadContext): GroupChildMediaData | undefined {
  if (el.name === "wps:wsp") return parseWpsChildMediaData(el, ctx);
  if (el.name === "pic:pic") {
    return parsePicChildMediaData(el, ctx) as GroupChildMediaData | undefined;
  }
  if (el.name === "wpg:grpSp") return parseNestedGroup(el, ctx);
  if (el.name === "wpg:graphicFrame") return parseGroupGraphicFrame(el, ctx);
  if (el.name === "wpg:contentPart" || el.name === "wp:contentPart") return parseContentPart(el);
  return undefined;
}

/**
 * Parse a wpg:graphicFrame group child (CT_GraphicFrame). Charts are the
 * payload Word produces in groups; the chart part is re-registered on
 * generate from the parsed chartOptions.
 */
function parseGroupGraphicFrame(el: Element, ctx: DocxReadContext): ChartMediaData | undefined {
  const chartRef = findFirst(el, "c:chart");
  if (!chartRef) return undefined;
  const rId = attr(chartRef, "r:id");
  const chartPath = rId ? lookupRId(ctx.docx.partRefs.charts, rId) : undefined;
  if (!chartPath) return undefined;

  const chartXml = ctx.docx.doc.get(chartPath);
  if (!chartXml) return undefined;
  // Full c:chartSpace model via the core descriptor (same descriptor that
  // stringifies on generate) — keeps axes, externalData, spPr, dLbls, …
  const chartOpts = chartSpaceDesc.parse(chartXml, ctx);
  if (!chartOpts.type) return undefined;

  const md: ChartMediaData = {
    type: "chart",
    transformation: readDirectXfrmTransformation(el),
    chartOptions: chartOpts as ChartSpaceOptions,
  };

  const cNvPr = findChild(el, "wpg:cNvPr") ?? findChild(el, "wp:cNvPr");
  if (cNvPr) {
    const nvp: NonVisualPropertiesOptions = {};
    const id = attrNum(cNvPr, "id");
    const name = attr(cNvPr, "name");
    const descr = attr(cNvPr, "descr");
    const title = attr(cNvPr, "title");
    if (id !== undefined) nvp.id = id;
    if (name) nvp.name = name;
    if (descr) nvp.description = descr;
    if (title) nvp.title = title;
    if (Object.keys(nvp).length > 0) md.nonVisualProperties = nvp;
  }
  const cNvFrPr = findChild(el, "wpg:cNvFrPr") ?? findChild(el, "wp:cNvFrPr");
  if (cNvFrPr) md.graphicFrameLocks = readGraphicFrameLocks(cNvFrPr);

  return md;
}

/**
 * Parse a wp:/wpg:contentPart element (CT_WordprocessingContentPart). The
 * r:id is captured verbatim — the relationship itself is not re-registered
 * on generate.
 */
function parseContentPart(el: Element): ContentPartMediaData {
  const md: ContentPartMediaData = {
    type: "contentPart",
    referenceId: attr(el, "r:id") ?? "",
    transformation: readDirectXfrmTransformation(el),
  };
  const bwMode = attr(el, "bwMode");
  if (bwMode) md.blackWhiteMode = bwMode as BlackWhiteMode;

  const nv = findChild(el, "wp:nvContentPartPr") ?? findChild(el, "wpg:nvContentPartPr");
  if (nv) {
    const cNvPr = findChild(nv, "wp:cNvPr") ?? findChild(nv, "wpg:cNvPr");
    const cpPr = findChild(nv, "wp:cNvContentPartPr") ?? findChild(nv, "wpg:cNvContentPartPr");
    const nvp: NonVisualContentPartNv = {};
    if (cNvPr) {
      const id = attrNum(cNvPr, "id");
      const name = attr(cNvPr, "name");
      const descr = attr(cNvPr, "descr");
      const title = attr(cNvPr, "title");
      if (id !== undefined) nvp.id = id;
      if (name) nvp.name = name;
      if (descr) nvp.description = descr;
      if (title) nvp.title = title;
    }
    const contentPart = parseNonVisualContentPartProperties(cpPr);
    if (contentPart) nvp.contentPart = contentPart;
    if (Object.keys(nvp).length > 0) md.nonVisualProperties = nvp;
  }
  return md;
}

/** Non-visual properties of a content part (cNvPr + cNvContentPartPr). */
interface NonVisualContentPartNv extends NonVisualPropertiesOptions {
  contentPart?: NonVisualContentPartPropertiesOptions;
}

/**
 * Parse a nested wpg:grpSp (CT_WordprocessingGroup) into a GroupMediaData child.
 * Mirrors the top-level group: grpSpPr transform + chOff/chExt + fill, with its
 * own children (which may nest further groups).
 */
function parseNestedGroup(grpSpEl: Element, ctx: DocxReadContext): GroupMediaData {
  const grpSpPr = findChild(grpSpEl, "wpg:grpSpPr");
  const { childOffset, childExtent } = readGroupCoords(grpSpPr);
  const result: GroupMediaData = {
    type: "wpg",
    transformation: readChildTransformation(grpSpPr),
    children: parseGroupChildren(grpSpEl, ctx),
  };
  if (childOffset) result.childOffset = childOffset;
  if (childExtent) result.childExtent = childExtent;
  const grpSpLocks = readGrpSpLocks(findChild(grpSpEl, "wpg:cNvGrpSpPr"));
  if (grpSpLocks) result.groupShapeLocks = grpSpLocks;
  if (grpSpPr) {
    const fill = readShapeFill(grpSpPr, ctx);
    if (fill) result.fill = fill;
  }
  return result;
}

// ── Floating (anchor) parse helpers ─────────────────────────────────────────

/** Map wp:positionH/V children + relativeFrom into a position-options object. */
function readPosition(
  posEl: Element,
): HorizontalPositionOptions | VerticalPositionOptions | undefined {
  const relative = attr(posEl, "relativeFrom");
  const alignEl = findChild(posEl, "wp:align");
  const posOffset = findChild(posEl, "wp:posOffset");
  const result: { relative?: string; align?: string; offset?: number } = {};
  if (relative) result.relative = relative;
  if (alignEl) {
    const a = textOf(alignEl);
    if (a) result.align = a;
  } else if (posOffset) {
    const val = Number(textOf(posOffset));
    if (!isNaN(val)) result.offset = val;
  }
  return Object.keys(result).length > 0
    ? (result as HorizontalPositionOptions | VerticalPositionOptions)
    : undefined;
}

/** Read wp:wrapPolygon (start + lineTo points) into a WrapPolygon, if present. */
function readWrapPolygon(el: Element): WrapPolygon | undefined {
  const poly = findChild(el, "wp:wrapPolygon");
  if (!poly) return undefined;
  const points: { x: number; y: number }[] = [];
  const start = findChild(poly, "wp:start");
  if (start) points.push({ x: attrNum(start, "x") ?? 0, y: attrNum(start, "y") ?? 0 });
  for (const child of poly.elements ?? []) {
    if (child.name === "wp:lineTo") {
      points.push({ x: attrNum(child, "x") ?? 0, y: attrNum(child, "y") ?? 0 });
    }
  }
  if (points.length === 0) return undefined;
  return { edited: attrBool(poly, "edited"), points };
}

/** Map the wp:anchor wrap child element into a TextWrapping ({ type, side? }). */
function readWrap(anchor: Element): TextWrapping | undefined {
  const WRAP_TYPE: ReadonlyArray<[string, TextWrapping["type"]]> = [
    ["wrapNone", TextWrappingType.NONE],
    ["wrapSquare", TextWrappingType.SQUARE],
    ["wrapTight", TextWrappingType.TIGHT],
    ["wrapTopAndBottom", TextWrappingType.TOP_AND_BOTTOM],
    ["wrapThrough", TextWrappingType.THROUGH],
  ];
  for (const [name, type] of WRAP_TYPE) {
    const el = findChild(anchor, `wp:${name}`);
    if (!el) continue;
    const wrap: TextWrapping = { type };
    const side = attr(el, "wrapText");
    if (side) wrap.side = side as TextWrapping["side"];
    // wrapTight/wrapThrough carry a contour polygon; preserve it verbatim.
    if (name === "wrapTight" || name === "wrapThrough") {
      const polygon = readWrapPolygon(el);
      if (polygon) wrap.polygon = polygon;
    }
    return wrap;
  }
  return undefined;
}

// ── Common helpers ──────────────────────────────────────────────────────────

function getDrawingExtent(el: Element): { width?: number; height?: number } {
  const inline = findFirst(el, "wp:inline");
  const anchor = inline ? undefined : findFirst(el, "wp:anchor");
  const parent = inline ?? anchor;
  if (!parent) return {};

  const extent = findChild(parent, "wp:extent");
  if (!extent) return {};

  const cxEmu = attrNum(extent, "cx");
  const cyEmu = attrNum(extent, "cy");
  return {
    ...(cxEmu !== undefined ? { width: cxEmu } : {}),
    ...(cyEmu !== undefined ? { height: cyEmu } : {}),
  };
}

// ── Chart parsing ───────────────────────────────────────────────────────────

/**
 * Look up a relationship ID in a map, with fallback for double "rId" prefix
 * that the library's generation code produces (e.g. "rIdrId7" → "rId7").
 */
function lookupRId(map: Map<string, string>, rId: string | undefined): string | undefined {
  if (!rId) return undefined;
  const direct = map.get(rId);
  if (direct) return direct;
  // Fallback: strip one "rId" prefix when the value starts with "rIdrId"
  if (rId.startsWith("rIdrId")) return map.get(rId.slice(3));
  return undefined;
}

function parseChartDrawing(el: Element, ctx: DocxReadContext): { chart: ChartOptions } | undefined {
  const chartRef = findFirst(el, "c:chart");
  if (!chartRef) return undefined;

  const rId = attr(chartRef, "r:id");
  const chartPath = lookupRId(ctx.docx.partRefs.charts, rId);
  if (!chartPath) return undefined;

  const chartXml = ctx.docx.doc.get(chartPath);
  if (!chartXml) return undefined;

  // Full c:chartSpace model via the core descriptor — the same descriptor
  // that stringifies on generate, so every field round-trips symmetrically.
  const chartSpace = chartSpaceDesc.parse(chartXml, ctx);
  if (!chartSpace.type) return undefined;

  // Anchor wrapper fields: extent, alt text, frame locks.
  const info = parseAnchorOrInline(el);
  const ext = getDrawingExtent(el);
  const opts: ChartOptions = {
    ...chartSpace,
    // wp:extent always carries both cx and cy in valid documents; fall back
    // to zeros so the required transformation field stays well-formed.
    transformation: { width: ext.width ?? 0, height: ext.height ?? 0 },
  };
  if (info?.graphicFrameLocks !== undefined) {
    opts.graphicFrameLocks = info.graphicFrameLocks;
  }

  return { chart: opts };
}

// ── SmartArt parsing ────────────────────────────────────────────────────────

function parseSmartArtDrawing(
  el: Element,
  ctx: DocxReadContext,
): { smartArt: SmartArtOptions } | undefined {
  const relIds = findFirst(el, "dgm:relIds");
  if (!relIds) return undefined;

  const rId = attr(relIds, "r:dm");
  const dataPath = lookupRId(ctx.docx.partRefs.diagramData, rId);
  if (!dataPath) return undefined;

  const dataEl = ctx.docx.doc.get(dataPath);
  if (!dataEl) return undefined;

  const opts = parseSmartArtDataXml(dataEl);
  if (!opts) return undefined;

  // Custom definitions come back structured; built-in stubs fold to their id
  // string so round-tripping a built-in diagram keeps the compact form.
  const layoutEl = readDiagramPart(ctx, relIds, "r:lo", ctx.docx.partRefs.diagramLayout);
  if (layoutEl) {
    const layout = parseLayoutDefinition(layoutEl);
    const id = layout.uniqueId?.split("/").pop();
    opts.layout = id && id in LAYOUT_CATEGORIES ? id : layout;
  }
  const styleEl = readDiagramPart(ctx, relIds, "r:qs", ctx.docx.partRefs.diagramQuickStyle);
  if (styleEl) {
    const style = parseStyleDefinition(styleEl);
    const id = style.uniqueId?.split("/").pop();
    opts.style = id && id in STYLE_CATEGORIES ? id : style;
  }
  const colorEl = readDiagramPart(ctx, relIds, "r:cs", ctx.docx.partRefs.diagramColors);
  if (colorEl) {
    const color = parseColorDefinition(colorEl);
    const id = color.uniqueId?.split("/").pop();
    opts.color = id && id in COLOR_CATEGORIES ? id : color;
  }

  const ext = getDrawingExtent(el);
  if (ext.width !== undefined || ext.height !== undefined) {
    (opts as Record<string, unknown>).transformation = {
      ...ext,
    };
  }

  return { smartArt: opts as unknown as SmartArtOptions };
}

/** Resolve a dgm:relIds attribute through a part-kind map to its element. */
function readDiagramPart(
  ctx: DocxReadContext,
  relIds: Element,
  attrName: "r:lo" | "r:qs" | "r:cs",
  refs: Map<string, string>,
): Element | undefined {
  const rId = attr(relIds, attrName);
  if (!rId) return undefined;
  const path = lookupRId(refs, rId);
  return path ? ctx.docx.doc.get(path) : undefined;
}

/**
 * Parse dgm:dataModel element into SmartArtOptions.
 */
function parseSmartArtDataXml(el: Element): Record<string, unknown> | undefined {
  const ptLst = findChild(el, "dgm:ptLst");
  if (!ptLst) return undefined;

  const opts: Record<string, unknown> = {};
  const nodeMap = new Map<string, string>(); // modelId → text

  for (const pt of ptLst.elements ?? []) {
    if (pt.name !== "dgm:pt") continue;
    const type = attr(pt, "type");
    const modelId = attr(pt, "modelId");

    if (type === "doc") {
      // Extract layout/style/color from prSet URIs
      const prSet = findChild(pt, "dgm:prSet");
      if (prSet) {
        const loTypeId = attr(prSet, "loTypeId") ?? "";
        const qsTypeId = attr(prSet, "qsTypeId") ?? "";
        const csTypeId = attr(prSet, "csTypeId") ?? "";

        const layout = loTypeId.split("/").pop();
        if (layout) opts.layout = layout;
        const style = qsTypeId.split("/").pop();
        if (style) opts.style = style;
        const color = csTypeId.split("/").pop();
        if (color) opts.color = color;
      }
    } else if (type === "node" && modelId) {
      // Extract text: dgm:t → a:p → a:r → a:t
      const t = findFirst(pt, "a:t");
      nodeMap.set(modelId, t ? (textOf(t) ?? "") : "");
    }
  }

  // Build tree from connections
  const cxnLst = findChild(el, "dgm:cxnLst");
  if (!cxnLst) {
    opts.nodes = [];
    return opts;
  }

  // Map: parentId → childIds
  const childrenMap = new Map<string, string[]>();
  for (const cxn of cxnLst.elements ?? []) {
    if (cxn.name !== "dgm:cxn") continue;
    const srcId = attr(cxn, "srcId");
    const destId = attr(cxn, "destId");
    if (!srcId || !destId || !nodeMap.has(destId)) continue;

    let arr = childrenMap.get(srcId);
    if (!arr) {
      arr = [];
      childrenMap.set(srcId, arr);
    }
    arr.push(destId);
  }

  // Root children are connected from doc node (modelId="0")
  const topIds = childrenMap.get("0") ?? [];
  opts.nodes = topIds.map((id) => buildSmartArtNode(id, nodeMap, childrenMap));

  return opts;
}

function buildSmartArtNode(
  id: string,
  nodeMap: Map<string, string>,
  childrenMap: Map<string, string[]>,
): { text: string; children?: unknown[] } {
  const text = nodeMap.get(id) ?? "";
  const childIds = childrenMap.get(id) ?? [];

  if (childIds.length === 0) return { text };
  return { text, children: childIds.map((cid) => buildSmartArtNode(cid, nodeMap, childrenMap)) };
}
