/**
 * Drawing parser for DOCX documents.
 *
 * Parses w:drawing elements and extracts image, chart, or SmartArt data.
 *
 * @module
 */
import {
  blipDesc,
  convertEmuToPixels,
  customGeometryDesc,
  effectListDesc,
  fillDesc,
  outlineDesc,
  parseColorChoice,
  presetGeometryDesc,
} from "@office-open/core";
import { attr, attrBool, attrNum, findChild, findDeep, textOf } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import type { ChartOptions } from "@parts/paragraph/run/chart-run";
import type { ImageOptions } from "@parts/paragraph/run/image-run";
import type { SmartArtOptions } from "@parts/paragraph/run/smartart-run";
import type { WpgGroupRunOptions } from "@parts/paragraph/run/wpg-group-run";
import type { WpsShapeRunOptions } from "@parts/paragraph/run/wps-shape-run";
import type {
  GroupChildMediaData,
  MediaData,
  MediaDataTransformation,
  WpgCommonMediaData,
  WpgMediaData,
  WpsMediaData,
} from "@shared/media";
import type { NonVisualPropertiesOptions } from "@shared/media/data";

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
import type { SourceRectangleOptions } from "./inline/graphic/graphic-data/pic/blip/source-rectangle";
import type { ChildOffset, ChildExtent } from "./inline/graphic/graphic-data/wpg/wpg-group";
import { parseBodyProperties } from "./inline/graphic/graphic-data/wps/body-properties";
import type { NonVisualShapePropertiesOptions } from "./inline/graphic/graphic-data/wps/non-visual-shape-properties";
import type {
  ShapeStyleOptions,
  StyleMatrixReferenceOptions,
  WpsShapeCoreOptions,
} from "./inline/graphic/graphic-data/wps/wps-shape";
import { TextWrappingType } from "./text-wrap";
import type { TextWrapping } from "./text-wrap";

/** Union type for parsed drawing child wrappers. */
export type DrawingChild =
  | { image: ImageOptions }
  | { chart: ChartOptions }
  | { smartArt: SmartArtOptions }
  | { wpsShape: WpsShapeRunOptions }
  | { wpgGroup: WpgGroupRunOptions };

/**
 * Parse a w:drawing element and dispatch to the correct parser
 * based on the graphicData URI.
 */
export function parseDrawingRun(el: Element, ctx: DocxReadContext): DrawingChild | undefined {
  const graphicData = findDeep(el, "a:graphicData")[0];
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
  return parseImageRun(el, ctx);
}

/**
 * Determine image type from file extension or MIME type.
 */
export function imageTypeFromPath(
  path: string,
): "jpg" | "png" | "gif" | "bmp" | "tif" | "ico" | "emf" | "wmf" {
  const ext = path.split(".").pop()?.toLowerCase() ?? "";
  switch (ext) {
    case "jpg":
    case "jpeg":
      return "jpg";
    case "png":
      return "png";
    case "gif":
      return "gif";
    case "bmp":
      return "bmp";
    case "tif":
    case "tiff":
      return "tif";
    case "ico":
      return "ico";
    case "emf":
      return "emf";
    case "wmf":
      return "wmf";
    default:
      return "png"; // fallback
  }
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
  if (a["noGrp"] !== undefined) result.noGrp = a["noGrp"] !== "0";
  if (a["noDrilldown"] !== undefined) result.noDrilldown = a["noDrilldown"] !== "0";
  if (a["noSelect"] !== undefined) result.noSelect = a["noSelect"] !== "0";
  if (a["noChangeAspect"] !== undefined) result.noChangeAspect = a["noChangeAspect"] !== "0";
  if (a["noMove"] !== undefined) result.noMove = a["noMove"] !== "0";
  if (a["noResize"] !== undefined) result.noResize = a["noResize"] !== "0";
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
  if (a["noGrp"] !== undefined) result.noGrp = a["noGrp"] !== "0";
  if (a["noUngrp"] !== undefined) result.noUngrp = a["noUngrp"] !== "0";
  if (a["noSelect"] !== undefined) result.noSelect = a["noSelect"] !== "0";
  if (a["noRot"] !== undefined) result.noRot = a["noRot"] !== "0";
  if (a["noChangeAspect"] !== undefined) result.noChangeAspect = a["noChangeAspect"] !== "0";
  if (a["noMove"] !== undefined) result.noMove = a["noMove"] !== "0";
  if (a["noResize"] !== undefined) result.noResize = a["noResize"] !== "0";
  return Object.keys(result).length === 0 ? undefined : (result as GroupShapeLocksOptions);
}

/**
 * Extract {@link AnchorInfo} from the drawing's wp:inline or wp:anchor.
 * Returns `null` when the drawing has neither wrapper.
 */
function parseAnchorOrInline(el: Element): AnchorInfo | null {
  const inline = findDeep(el, "wp:inline")[0];
  const anchor = inline ? undefined : findDeep(el, "wp:anchor")[0];
  const parent = inline ?? anchor;
  if (!parent) return null;

  const info: AnchorInfo = {};

  // Extent (EMU → pixels)
  const extent = findChild(parent, "wp:extent");
  if (extent) {
    const cxEmu = attrNum(extent, "cx");
    const cyEmu = attrNum(extent, "cy");
    if (cxEmu !== undefined) info.width = convertEmuToPixels(cxEmu);
    if (cyEmu !== undefined) info.height = convertEmuToPixels(cyEmu);
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
    const id = attr(docPr, "id");
    const name = attr(docPr, "name");
    const descr = attr(docPr, "descr");
    const title = attr(docPr, "title");
    if (id !== undefined || name || descr || title) {
      const alt: Partial<DocPropertiesOptions> = {};
      if (id !== undefined) alt.id = id;
      if (name) alt.name = name;
      if (descr) alt.description = descr;
      if (title) alt.title = title;
      info.altText = alt as DocPropertiesOptions;
    }
  }

  // Graphic frame locks (wp:cNvGraphicFramePr) — preserved verbatim.
  const cNvGraphicFramePr = findChild(parent, "wp:cNvGraphicFramePr");
  if (cNvGraphicFramePr) info.graphicFrameLocks = readGraphicFrameLocks(cNvGraphicFramePr);

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
 * Parse a w:drawing element and return image data wrapped in { image: ... }.
 */
export function parseImageRun(
  el: Element,
  ctx: DocxReadContext,
): { image: ImageOptions } | undefined {
  const info = parseAnchorOrInline(el);
  if (!info) return undefined;

  // Get graphic → graphicData → blip
  const blip = findDeep(el, "a:blip")[0];
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
  const blipFill = findDeep(el, "pic:blipFill")[0];
  if (blipFill) {
    const srcRect = readSourceRectangle(blipFill);
    if (srcRect) imageOpts.sourceRectangle = srcRect;
  }

  // Picture non-visual properties (pic:nvPicPr/pic:cNvPr)
  const cNvPr = readPicCnvPr(el);
  if (cNvPr) imageOpts.nonVisualProperties = cNvPr;

  // Picture shape properties (pic:spPr): outline + fill + effects round-trip
  // via the shared core descriptors (bidirectional).
  const picSpPr = findDeep(el, "pic:spPr")[0];
  if (picSpPr) {
    const fill = readShapeFill(picSpPr, ctx);
    if (fill) imageOpts.fill = fill;
    const ln = findChild(picSpPr, "a:ln");
    if (ln) imageOpts.outline = outlineDesc.parse(ln, ctx);
    const effectLst = findChild(picSpPr, "a:effectLst");
    if (effectLst) imageOpts.effects = effectListDesc.parse(effectLst, ctx);
  }

  // Blip recolor effects (a:lum/a:hsl/a:tint/...) under a:blip — image
  // brightness/contrast/tint adjustments applied directly to the image data.
  const blipResult = blipDesc.parse(blip, ctx);
  if (blipResult.blipEffects) imageOpts.blipEffects = blipResult.blipEffects;

  // Blip extension: a14:useLocalDpi (rendering hint, round-trip verbatim).
  const useLocalDpi = readBlipUseLocalDpi(blip);
  if (useLocalDpi !== undefined) imageOpts.useLocalDpi = useLocalDpi;

  return { image: imageOpts as unknown as ImageOptions };
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
      return val !== "0";
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
function readPicCnvPr(el: Element): NonVisualPropertiesOptions | undefined {
  const nvPicPr = findDeep(el, "pic:nvPicPr")[0];
  if (!nvPicPr) return undefined;
  const result: NonVisualPropertiesOptions = {};
  const cNvPr = findChild(nvPicPr, "pic:cNvPr");
  if (cNvPr) {
    const id = attrNum(cNvPr, "id");
    const name = attr(cNvPr, "name");
    const descr = attr(cNvPr, "descr");
    if (id !== undefined) result.id = id;
    if (name) result.name = name;
    if (descr) result.description = descr;
  }
  // pic:cNvPicPr sibling — only preferRelativeResize is tracked (Word omits
  // the default true; an explicit false round-trips as "0").
  const cNvPicPr = findChild(nvPicPr, "pic:cNvPicPr");
  if (cNvPicPr) {
    const preferRelativeResize = attrBool(cNvPicPr, "preferRelativeResize");
    if (preferRelativeResize !== undefined) result.preferRelativeResize = preferRelativeResize;
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
    findChild(parent, "a:grpFill");
  if (!fillChild) return undefined;
  return fillDesc.parse(parent, ctx);
}

/**
 * Read a single style-matrix reference (a:lnRef/a:fillRef/a:effectRef/a:fontRef):
 * the `idx` attribute plus an optional EG_ColorChoice color override.
 */
function parseStyleRef(el: Element, ctx: DocxReadContext): StyleMatrixReferenceOptions | undefined {
  const idx = attr(el, "idx");
  if (idx === undefined) return undefined;
  const result: StyleMatrixReferenceOptions = { idx };
  const color = parseColorChoice(el, ctx);
  if (color && Object.keys(color).length > 0) result.color = color;
  return result as StyleMatrixReferenceOptions;
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
function parseWpsShapeCore(wspEl: Element, ctx: DocxReadContext): WpsShapeCoreOptions {
  const result: Partial<WpsShapeCoreOptions> = {};

  // Text content — w:txbxContent (w namespace, per CT_TxbxContent → w:EG_BlockLevelElts)
  // holds the shape's paragraphs, even when wrapped in wps:txbx.
  const txbxContent = findDeep(wspEl, "w:txbxContent")[0];
  const children: WpsShapeCoreOptions["children"] = [];
  if (txbxContent) {
    for (const child of txbxContent.elements ?? []) {
      if (child.name === "w:p") children.push(parseParagraph(child, ctx));
    }
  }
  result.children = children;

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
  }

  // Body properties (wps:bodyPr)
  const bodyPr = findChild(wspEl, "wps:bodyPr");
  if (bodyPr) result.bodyProperties = parseBodyProperties(bodyPr);

  // Shape style (wps:style) — theme references (lnRef/fillRef/effectRef/fontRef)
  const styleEl = findChild(wspEl, "wps:style");
  if (styleEl) result.style = parseShapeStyle(styleEl, ctx);

  return result as WpsShapeCoreOptions;
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
 * Parse a `wps:wsp` nested inside a wpg group into a {@link WpsMediaData} child
 * (its transformation kept as EMU via {@link readChildTransformation}).
 */
function parseWpsChildMediaData(wspEl: Element, ctx: DocxReadContext): WpsMediaData | undefined {
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
  const blip = findDeep(picEl, "a:blip")[0];
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
  const cNvPr = readPicCnvPr(picEl);
  if (cNvPr) result.nonVisualProperties = cNvPr;
  // Grouped picture spPr (fill/outline) rides on WpgCommonMediaData so it
  // round-trips through stringifyGroupChild → stringifyShapeProps.
  if (spPr) {
    const fill = readShapeFill(spPr, ctx);
    if (fill) (result as MediaData & WpgCommonMediaData).fill = fill;
    const ln = findChild(spPr, "a:ln");
    if (ln) (result as MediaData & WpgCommonMediaData).outline = outlineDesc.parse(ln, ctx);
  }
  return result;
}

/**
 * Parse a standalone wps shape drawing (graphicData URI wordprocessingShape).
 */
function parseWpsShapeDrawing(
  el: Element,
  ctx: DocxReadContext,
): { wpsShape: WpsShapeRunOptions } | undefined {
  const wsp = findDeep(el, "wps:wsp")[0];
  if (!wsp) return undefined;

  const info = parseAnchorOrInline(el) ?? {};
  const data = parseWpsShapeCore(wsp, ctx);

  const shape: WpsShapeRunOptions = {
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

  return { wpsShape: shape as WpsShapeRunOptions };
}

/**
 * Parse a wpg group drawing (graphicData URI wordprocessingGroup).
 */
function parseWpgGroupDrawing(
  el: Element,
  ctx: DocxReadContext,
): { wpgGroup: WpgGroupRunOptions } | undefined {
  const wgp = findDeep(el, "wpg:wgp")[0];
  if (!wgp) return undefined;

  const info = parseAnchorOrInline(el) ?? {};
  const grpSpPr = findChild(wgp, "wpg:grpSpPr");
  const { childOffset, childExtent } = readGroupCoords(grpSpPr);

  const group: WpgGroupRunOptions = {
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

  return { wpgGroup: group as WpgGroupRunOptions };
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
  return undefined;
}

/**
 * Parse a nested wpg:grpSp (CT_WordprocessingGroup) into a WpgMediaData child.
 * Mirrors the top-level group: grpSpPr transform + chOff/chExt + fill, with its
 * own children (which may nest further groups).
 */
function parseNestedGroup(grpSpEl: Element, ctx: DocxReadContext): WpgMediaData {
  const grpSpPr = findChild(grpSpEl, "wpg:grpSpPr");
  const { childOffset, childExtent } = readGroupCoords(grpSpPr);
  const result: WpgMediaData = {
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
    return wrap;
  }
  return undefined;
}

// ── Common helpers ──────────────────────────────────────────────────────────

function getDrawingExtent(el: Element): { width?: number; height?: number } {
  const inline = findDeep(el, "wp:inline")[0];
  const anchor = inline ? undefined : findDeep(el, "wp:anchor")[0];
  const parent = inline ?? anchor;
  if (!parent) return {};

  const extent = findChild(parent, "wp:extent");
  if (!extent) return {};

  const cxEmu = attrNum(extent, "cx");
  const cyEmu = attrNum(extent, "cy");
  return {
    ...(cxEmu !== undefined ? { width: convertEmuToPixels(cxEmu) } : {}),
    ...(cyEmu !== undefined ? { height: convertEmuToPixels(cyEmu) } : {}),
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
  const chartRef = findDeep(el, "c:chart")[0];
  if (!chartRef) return undefined;

  const rId = attr(chartRef, "r:id");
  const chartPath = lookupRId(ctx.docx.partRefs.charts, rId);
  if (!chartPath) return undefined;

  const chartXml = ctx.docx.doc.get(chartPath);
  if (!chartXml) return undefined;

  const opts = parseChartXml(chartXml);
  if (!opts) return undefined;

  const ext = getDrawingExtent(el);
  if (ext.width !== undefined || ext.height !== undefined) {
    (opts as Record<string, unknown>).transformation = {
      ...ext,
    };
  }

  return { chart: opts as unknown as ChartOptions };
}

/**
 * Parse c:chartSpace element into ChartOptions.
 */
function parseChartXml(el: Element): Record<string, unknown> | undefined {
  const chart = findChild(el, "c:chart");
  if (!chart) return undefined;

  const opts: Record<string, unknown> = {};

  // Title: c:chart → c:title → c:tx → c:rich → a:p → a:r → a:t
  const titleEl = findChild(chart, "c:title");
  if (titleEl) {
    const rich = findDeep(titleEl, "c:rich")[0];
    if (rich) {
      const t = findDeep(rich, "a:t")[0];
      if (t) {
        const title = textOf(t);
        if (title) opts.title = title;
      }
    }
  }

  // Plot area → chart type
  const plotArea = findChild(chart, "c:plotArea");
  if (!plotArea) return undefined;

  let chartType: string | undefined;
  let typeElement: Element | undefined;

  for (const child of plotArea.elements ?? []) {
    switch (child.name) {
      case "c:barChart": {
        const barDir = findChild(child, "c:barDir");
        chartType = barDir && attr(barDir, "val") === "bar" ? "bar" : "column";
        typeElement = child;
        break;
      }
      case "c:lineChart":
        chartType = "line";
        typeElement = child;
        break;
      case "c:pieChart":
        chartType = "pie";
        typeElement = child;
        break;
      case "c:areaChart":
        chartType = "area";
        typeElement = child;
        break;
      case "c:scatterChart":
        chartType = "scatter";
        typeElement = child;
        break;
    }
    if (chartType) break;
  }

  if (!chartType || !typeElement) return undefined;
  opts.type = chartType;

  // Parse series
  const series: { name: string; values: number[] }[] = [];
  let categories: string[] | undefined;

  for (const serEl of typeElement.elements ?? []) {
    if (serEl.name !== "c:ser") continue;

    // Series name: c:tx → c:strCache → c:pt → c:v
    const nameParts = extractStrCache(serEl, "c:tx");
    // Categories: c:cat → c:strCache → c:pt → c:v
    const cats = extractStrCache(serEl, "c:cat");
    if (cats.length > 0 && !categories) categories = cats;
    // Values: c:val → c:numCache → c:pt → c:v
    const vals = extractNumCache(serEl);

    series.push({ name: nameParts[0] ?? "", values: vals });
  }

  opts.categories = categories ?? [];
  opts.series = series;

  // Legend
  opts.showLegend = findChild(chart, "c:legend") !== undefined;

  // Style
  const styleEl = findChild(el, "c:style");
  if (styleEl) {
    const val = attrNum(styleEl, "val");
    if (val !== undefined) opts.style = val;
  }

  return opts;
}

/**
 * Extract string values from c:strCache within a container element.
 */
function extractStrCache(parent: Element, containerName: string): string[] {
  const container = findChild(parent, containerName);
  if (!container) return [];
  const cache = findDeep(container, "c:strCache")[0];
  if (!cache) return [];

  const values: string[] = [];
  for (const pt of cache.elements ?? []) {
    if (pt.name !== "c:pt") continue;
    const v = findChild(pt, "c:v");
    if (v) values.push(textOf(v) ?? "");
  }
  return values;
}

/**
 * Extract numeric values from c:numCache within a c:val container.
 */
function extractNumCache(parent: Element): number[] {
  const valEl = findChild(parent, "c:val");
  if (!valEl) return [];
  const cache = findDeep(valEl, "c:numCache")[0];
  if (!cache) return [];

  const values: number[] = [];
  for (const pt of cache.elements ?? []) {
    if (pt.name !== "c:pt") continue;
    const v = findChild(pt, "c:v");
    if (v) {
      const num = Number(textOf(v));
      if (!isNaN(num)) values.push(num);
    }
  }
  return values;
}

// ── SmartArt parsing ────────────────────────────────────────────────────────

function parseSmartArtDrawing(
  el: Element,
  ctx: DocxReadContext,
): { smartArt: SmartArtOptions } | undefined {
  const relIds = findDeep(el, "dgm:relIds")[0];
  if (!relIds) return undefined;

  const rId = attr(relIds, "r:dm");
  const dataPath = lookupRId(ctx.docx.partRefs.diagramData, rId);
  if (!dataPath) return undefined;

  const dataEl = ctx.docx.doc.get(dataPath);
  if (!dataEl) return undefined;

  const opts = parseSmartArtDataXml(dataEl);
  if (!opts) return undefined;

  const ext = getDrawingExtent(el);
  if (ext.width !== undefined || ext.height !== undefined) {
    (opts as Record<string, unknown>).transformation = {
      ...ext,
    };
  }

  return { smartArt: opts as unknown as SmartArtOptions };
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
      const t = findDeep(pt, "a:t")[0];
      nodeMap.set(modelId, t ? (textOf(t) ?? "") : "");
    }
  }

  // Build tree from connections
  const cxnLst = findChild(el, "dgm:cxnLst");
  if (!cxnLst) {
    opts.data = { nodes: [] };
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
  opts.data = { nodes: topIds.map((id) => buildSmartArtNode(id, nodeMap, childrenMap)) };

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
