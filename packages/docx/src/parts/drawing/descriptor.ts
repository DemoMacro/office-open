/**
 * Drawing descriptor for DOCX documents.
 *
 * Produces `<w:drawing>` XML directly from media data and options,
 * eliminating the Drawing/Inline/Anchor/Graphic/GraphicData/Pic XmlComponent
 * class chain (~10 instances per drawing in the old path).
 *
 * Common path (inline image/chart/smartart without advanced properties):
 * zero XmlComponent instances — pure string concatenation.
 *
 * Advanced properties (outline, fill, effects on images) and floating
 * positioning use core `create*()` + `.toXml({stack:[]})` for sub-elements
 * (lightweight BuilderElement instances, not deep hierarchies).
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_Drawing
 *
 * @module
 */

import { TargetModeType } from "@office-open/core";
import { convertToEmu, uniqueNumericIdCreator, uniqueId } from "@office-open/core";
import type { CustomDescriptor, WriteContext } from "@office-open/core/descriptor";
import type { FillOptions } from "@office-open/core/drawingml";
import {
  calculateEffectExtent,
  createColorElement,
  createEffectDag,
  customGeometryDesc,
  effectListDesc,
  fillDesc,
  outlineDesc,
  presetGeometryDesc,
  scene3DDesc,
  shape3DDesc,
  transform2DDesc,
} from "@office-open/core/drawingml";
import { escapeXml } from "@office-open/xml";
import { stringifyParagraphInline } from "@parts/inline";
import type { ParagraphOptions } from "@parts/paragraph/paragraph";
import type {
  ChartMediaData,
  ExtendedMediaData,
  GroupChildMediaData,
  MediaData,
  MediaDataTransformation,
  SmartArtMediaData,
  WpgMediaData,
  WpsMediaData,
} from "@shared/media";
import type { NonVisualPropertiesOptions } from "@shared/media/data";

import type { BodyContext, DocxReadContext } from "../../context";
import type { DocPropertiesOptions, HyperlinkOptions } from "./doc-properties/doc-properties";
// Import parse function from drawing-parse.ts (parse path)
import { parseDrawingRun } from "./drawing-parse";
import type { Floating, HorizontalPositionOptions, VerticalPositionOptions } from "./floating";
import type { Margins } from "./floating";
import { HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom } from "./floating";
import type { BlipEffectsOptions } from "./inline/graphic/graphic-data/pic/blip/blip-effects";
import type { SourceRectangleOptions } from "./inline/graphic/graphic-data/pic/blip/source-rectangle";
import type { TileOptions } from "./inline/graphic/graphic-data/pic/blip/tile";
import type { EffectListOptions } from "./inline/graphic/graphic-data/pic/effects/effect-list";
import type { OutlineOptions } from "./inline/graphic/graphic-data/pic/outline/outline";
import type { ChildOffset, ChildExtent } from "./inline/graphic/graphic-data/wpg/wpg-group";
// wpg/wps types only
import {
  createBodyProperties,
  type BodyPropertiesOptions,
} from "./inline/graphic/graphic-data/wps/body-properties";
import type { NonVisualShapePropertiesOptions } from "./inline/graphic/graphic-data/wps/non-visual-shape-properties";
import type {
  ShapeStyleOptions,
  StyleMatrixReferenceOptions,
  WpsShapeCoreOptions,
} from "./inline/graphic/graphic-data/wps/wps-shape";
import { TextWrappingSide, TextWrappingType } from "./text-wrap";
import type { TextWrapping } from "./text-wrap";

// Noop context for drawingml descriptors that don't use WriteContext
const NOOP_CTX: WriteContext = {
  addRelationship: () => "",
  addMedia: () => "",
};

// ── Options ──

/**
 * Options for the drawing descriptor.
 *
 * Combines media data with optional visual properties.
 */

/** Locking flags for wp:cNvGraphicFramePr (CT_GraphicalObjectFrameLocking). */
export interface GraphicFrameLocksOptions {
  noGrp?: boolean;
  noDrilldown?: boolean;
  noSelect?: boolean;
  noChangeAspect?: boolean;
  noMove?: boolean;
  noResize?: boolean;
}

/**
 * Group shape locks (CT_GroupLocking) carried inside wpg:cNvGrpSpPr.
 * Distinct from GraphicFrameLocksOptions: groups use noUngrp/noRot instead of noDrilldown.
 */
export interface GroupShapeLocksOptions {
  noGrp?: boolean;
  noUngrp?: boolean;
  noSelect?: boolean;
  noRot?: boolean;
  noChangeAspect?: boolean;
  noMove?: boolean;
  noResize?: boolean;
}

export interface DrawingDescriptorOptions {
  /** Media data (image, chart, smartart, wps, wpg) */
  mediaData: ExtendedMediaData;
  /** Non-visual document properties (name, description, hyperlinks) */
  docProperties?: DocPropertiesOptions;
  /** Floating/anchored positioning (omit for inline) */
  floating?: Floating;
  /** Shape outline */
  outline?: OutlineOptions;
  /** Shape fill */
  fill?: FillOptions;
  /** Shape effects (shadow, glow, etc.) */
  effects?: EffectListOptions;
  /** Image blip effects (brightness, contrast, etc.) */
  blipEffects?: BlipEffectsOptions;
  /** Image tile fill mode */
  tile?: TileOptions;
  /** Graphic frame locks (wp:cNvGraphicFramePr). `{}` → empty element; omit → authoring default. */
  graphicFrameLocks?: GraphicFrameLocksOptions | null;
}

// ── ID generation ──

let _docPropsIdGen = uniqueNumericIdCreator();

/** Reset the doc properties ID generator (for testing). */
export const resetDrawingIdGen = (): void => {
  _docPropsIdGen = uniqueNumericIdCreator();
};

// ── Constants ──

const GRAPHIC_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const PIC_URI = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const DGM_URI = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const WPS_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const WPG_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
const HYPERLINK_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
// Blip extension URIs (a:extLst under a:blip).
const SVG_BLIP_EXT_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
const USE_LOCAL_DPI_EXT_URI = "{28A0092B-C50C-407E-A947-70E740481C1C}";
const A14_NS = "http://schemas.microsoft.com/office/drawing/2010/main";

/**
 * Build the `a14:useLocalDpi` blip extension. Returns "" when the hint is
 * absent (undefined) — Word's default. val="0" (useLocalDpi=false) is the
 * common Word emission; val="1" only when explicitly set.
 */
function buildUseLocalDpiExt(useLocalDpi?: boolean): string {
  if (useLocalDpi === undefined) return "";
  return `<a:ext uri="${USE_LOCAL_DPI_EXT_URI}"><a14:useLocalDpi xmlns:a14="${A14_NS}" val="${
    useLocalDpi ? "1" : "0"
  }"/></a:ext>`;
}

// ── Hyperlink handling ──

interface HyperlinkIds {
  clickId?: string;
  hoverId?: string;
}

function registerHyperlinks(
  hyperlink: HyperlinkOptions | undefined,
  ctx: BodyContext,
): HyperlinkIds {
  if (!hyperlink) return {};
  const result: HyperlinkIds = {};
  if (hyperlink.click) {
    const linkId = uniqueId();
    ctx.viewWrapper.relationships.addRelationship(
      linkId,
      HYPERLINK_REL,
      hyperlink.click,
      TargetModeType.EXTERNAL,
    );
    result.clickId = `rId${linkId}`;
  }
  if (hyperlink.hover) {
    const linkId = uniqueId();
    ctx.viewWrapper.relationships.addRelationship(
      linkId,
      HYPERLINK_REL,
      hyperlink.hover,
      TargetModeType.EXTERNAL,
    );
    result.hoverId = `rId${linkId}`;
  }
  return result;
}

function buildHyperlinkChildren(ids: HyperlinkIds): string {
  const parts: string[] = [];
  const aNs = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
  if (ids.clickId) parts.push(`<a:hlinkClick r:id="${ids.clickId}" ${aNs}/>`);
  if (ids.hoverId) parts.push(`<a:hlinkHover r:id="${ids.hoverId}" ${aNs}/>`);
  return parts.join("");
}

// ── DocPr ──

function stringifyDocPr(opts: DocPropertiesOptions | undefined, hlIds: HyperlinkIds): string {
  const id = opts?.id ?? _docPropsIdGen();
  const name = opts?.name ?? "";
  const attrs: string[] = [`id="${id}"`, `name="${escapeXml(name)}"`];
  if (opts?.description != null && opts.description !== undefined) {
    attrs.push(`descr="${escapeXml(opts.description)}"`);
  }
  if (opts?.title != null && opts.title !== undefined) {
    attrs.push(`title="${escapeXml(opts.title)}"`);
  }
  const hlXml = buildHyperlinkChildren(hlIds);
  if (hlXml) {
    return `<wp:docPr ${attrs.join(" ")}>${hlXml}</wp:docPr>`;
  }
  return `<wp:docPr ${attrs.join(" ")}/>`;
}

// ── BlipFill (image data reference) ──

/** Build `<a:srcRect .../>` from a crop spec, or "" when there is no crop. */
function buildSrcRectXml(srcRect: SourceRectangleOptions | undefined): string {
  if (!srcRect) return "";
  const srAttrs: string[] = [];
  if (srcRect.left !== undefined) srAttrs.push(`l="${srcRect.left}"`);
  if (srcRect.top !== undefined) srAttrs.push(`t="${srcRect.top}"`);
  if (srcRect.right !== undefined) srAttrs.push(`r="${srcRect.right}"`);
  if (srcRect.bottom !== undefined) srAttrs.push(`b="${srcRect.bottom}"`);
  return srAttrs.length ? `<a:srcRect ${srAttrs.join(" ")}/>` : "<a:srcRect/>";
}

function stringifyBlipFill(
  mediaData: MediaData,
  blipEffects?: BlipEffectsOptions,
  tile?: TileOptions,
): string {
  const fileName =
    mediaData.type === "svg" && "fallback" in mediaData
      ? mediaData.fallback.fileName
      : mediaData.fileName;

  const parts: string[] = [];

  // a:blip — cstate omitted unless set; Word's default is "none", so emitting
  // it unconditionally inflates round-trip output that originally had none.
  const blipAttrs: string[] = [`r:embed="{${escapeXml(fileName)}}"`];

  // Blip extension list: useLocalDpi (rendering hint) + SVG blip reference.
  // Both live in a single shared a:extLst; emitted only when at least one ext
  // is present so blips without extensions stay self-closing.
  const extParts: string[] = [];
  const useLocalDpiExt = buildUseLocalDpiExt(mediaData.useLocalDpi);
  if (useLocalDpiExt) extParts.push(useLocalDpiExt);
  if (mediaData.type === "svg") {
    extParts.push(
      `<a:ext uri="${SVG_BLIP_EXT_URI}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="{${escapeXml(
        mediaData.fileName,
      )}}"/></a:ext>`,
    );
  }
  const extLstXml = extParts.length > 0 ? `<a:extLst>${extParts.join("")}</a:extLst>` : "";

  // Blip effects
  const blipEffectsXml = blipEffects ? buildBlipEffectsXml(blipEffects) : "";

  const blipContent = extLstXml + blipEffectsXml;
  if (blipContent) {
    parts.push(`<a:blip ${blipAttrs.join(" ")}>${blipContent}</a:blip>`);
  } else {
    parts.push(`<a:blip ${blipAttrs.join(" ")}/>`);
  }

  // Source rectangle (blip crop)
  const srcRectXml = buildSrcRectXml(mediaData.sourceRectangle);
  if (srcRectXml) parts.push(srcRectXml);

  // Tile or stretch
  if (tile) {
    const tileAttrs: string[] = [];
    if (tile.tx !== undefined) tileAttrs.push(`tx="${tile.tx}"`);
    if (tile.ty !== undefined) tileAttrs.push(`ty="${tile.ty}"`);
    if (tile.sx !== undefined) tileAttrs.push(`sx="${tile.sx}"`);
    if (tile.sy !== undefined) tileAttrs.push(`sy="${tile.sy}"`);
    const tileAttrStr = tileAttrs.length ? " " + tileAttrs.join(" ") : "";
    parts.push(`<a:tile${tileAttrStr}/>`);
  } else {
    parts.push("<a:stretch><a:fillRect/></a:stretch>");
  }

  return `<pic:blipFill>${parts.join("")}</pic:blipFill>`;
}

function buildBlipEffectsXml(opts: BlipEffectsOptions): string {
  const parts: string[] = [];
  if (opts.grayscale) parts.push("<a:grayscl/>");
  if (opts.luminance) {
    const a: string[] = [];
    if (opts.luminance.bright !== undefined) a.push(`bright="${opts.luminance.bright}"`);
    if (opts.luminance.contrast !== undefined) a.push(`contrast="${opts.luminance.contrast}"`);
    parts.push(`<a:lum${a.length ? " " + a.join(" ") : ""}/>`);
  }
  if (opts.biLevel) parts.push(`<a:biLevel thresh="${opts.biLevel.threshold}"/>`);
  if (opts.blur) {
    const a: string[] = [];
    if (opts.blur.radius !== undefined) a.push(`rad="${opts.blur.radius}"`);
    if (opts.blur.grow === false) a.push('grow="0"');
    parts.push(`<a:blur${a.length ? " " + a.join(" ") : ""}/>`);
  }
  return parts.join("");
}

// ── Shape Properties (pic:spPr) ──

function stringifyShapeProps(
  transform: MediaDataTransformation,
  outline?: OutlineOptions,
  fill?: FillOptions,
  effects?: EffectListOptions,
): string {
  const parts: string[] = [];

  // Transform
  parts.push(
    transform2DDesc.stringify(
      {
        x: transform.offset?.emus?.x ?? 0,
        y: transform.offset?.emus?.y ?? 0,
        width: transform.emus.x,
        height: transform.emus.y,
        flipHorizontal: transform.flip?.horizontal,
        flipVertical: transform.flip?.vertical,
        rotation: transform.rotation,
      },
      NOOP_CTX,
    ) ?? "",
  );

  // Geometry (always rect — preset geometry variations not used for pic)
  parts.push('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>');

  if (fill) parts.push(fillDesc.stringify(fill, NOOP_CTX) ?? "");
  if (outline) parts.push(outlineDesc.stringify(outline, NOOP_CTX) ?? "");
  if (effects) parts.push(effectListDesc.stringify(effects, NOOP_CTX) ?? "");

  return `<pic:spPr bwMode="auto">${parts.join("")}</pic:spPr>`;
}

// ── Non-visual picture properties (pic:nvPicPr) ──

function stringifyNvPicPr(hlIds: HyperlinkIds, cNvPr?: NonVisualPropertiesOptions): string {
  const hlXml = buildHyperlinkChildren(hlIds);
  const id = cNvPr?.id ?? 0;
  const name = escapeXml(cNvPr?.name ?? "");
  // descr omitted when absent — Word never writes an empty descr attribute.
  const descrAttr = cNvPr?.description ? ` descr="${escapeXml(cNvPr.description)}"` : "";
  // preferRelativeResize defaults to true; only an explicit false is written.
  const cNvPicPrAttr = cNvPr?.preferRelativeResize === false ? ' preferRelativeResize="0"' : "";
  const cNvPrClose = hlXml ? `>${hlXml}</pic:cNvPr>` : "/>";
  return (
    `<pic:nvPicPr><pic:cNvPr id="${id}" name="${name}"${descrAttr}${cNvPrClose}` +
    `<pic:cNvPicPr${cNvPicPrAttr}><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr>`
  );
}

// ── WPS shape (pure string, no class instances) ──

function stringifyGroupTransform2D(
  transform: MediaDataTransformation,
  childOffset?: ChildOffset,
  childExtent?: ChildExtent,
): string {
  const attrs: string[] = [];
  if (transform.flip?.horizontal !== undefined) attrs.push(`flipH="${transform.flip.horizontal}"`);
  if (transform.flip?.vertical !== undefined) attrs.push(`flipV="${transform.flip.vertical}"`);
  if (transform.rotation !== undefined) attrs.push(`rot="${transform.rotation}"`);
  const attrStr = attrs.length ? " " + attrs.join(" ") : "";

  const off = `<a:off x="${transform.offset?.emus?.x ?? 0}" y="${transform.offset?.emus?.y ?? 0}"/>`;
  const ext = `<a:ext cx="${transform.emus.x}" cy="${transform.emus.y}"/>`;
  const childOffsetXml = childOffset ? `<a:chOff x="${childOffset.x}" y="${childOffset.y}"/>` : "";
  const childExtentXml = childExtent
    ? `<a:chExt cx="${childExtent.cx}" cy="${childExtent.cy}"/>`
    : "";

  return `<a:xfrm${attrStr}>${off}${ext}${childOffsetXml}${childExtentXml}</a:xfrm>`;
}

/** WpsShape options for stringification (extends WpsShapeCoreOptions with transformation). */
interface WpsStringifyOptions extends WpsShapeCoreOptions {
  transformation: MediaDataTransformation;
}

function stringifyWpsShape(opts: WpsStringifyOptions, ctx: BodyContext): string {
  const transform = opts.transformation;
  const spPrParts: string[] = [];
  spPrParts.push(
    transform2DDesc.stringify(
      {
        x: transform.offset?.emus?.x ?? 0,
        y: transform.offset?.emus?.y ?? 0,
        width: transform.emus.x,
        height: transform.emus.y,
        flipHorizontal: transform.flip?.horizontal,
        flipVertical: transform.flip?.vertical,
        rotation: transform.rotation,
      },
      NOOP_CTX,
    ) ?? "",
  );
  if (opts.customGeometry) {
    spPrParts.push(customGeometryDesc.stringify(opts.customGeometry, NOOP_CTX) ?? "");
  } else if (opts.presetGeometry) {
    spPrParts.push(presetGeometryDesc.stringify(opts.presetGeometry, NOOP_CTX) ?? "");
  } else {
    spPrParts.push('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>');
  }
  if (opts.fill) spPrParts.push(fillDesc.stringify(opts.fill, ctx) ?? "");
  if (opts.outline) spPrParts.push(outlineDesc.stringify(opts.outline, NOOP_CTX) ?? "");
  if (opts.effectDag) {
    spPrParts.push(createEffectDag(opts.effectDag));
  } else if (opts.effects) {
    spPrParts.push(effectListDesc.stringify(opts.effects, NOOP_CTX) ?? "");
  }
  if (opts.scene3d) spPrParts.push(scene3DDesc.stringify(opts.scene3d, NOOP_CTX) ?? "");
  if (opts.shape3d) spPrParts.push(shape3DDesc.stringify(opts.shape3d, NOOP_CTX) ?? "");

  // Non-visual shape properties — default txBox="1"
  const cNvSpPr = opts.nonVisualProperties
    ? stringifyNonVisualShapeProperties(opts.nonVisualProperties)
    : '<wps:cNvSpPr txBox="1"/>';

  // Paragraph children — pure JSON stringification
  const childXml =
    opts.children
      ?.map((c) => stringifyParagraphInline(c as ParagraphOptions | string, ctx))
      .join("") ?? "";

  // Shape style (wps:style) — theme references, emitted after spPr (XSD order)
  const styleXml = opts.style ? stringifyShapeStyle(opts.style) : "";
  // wps:txbx — only emit when the shape carries text (text boxes). Pure
  // geometry shapes (no paragraphs) omit txbx in the source.
  const txbxXml = childXml ? `<wps:txbx><w:txbxContent>${childXml}</w:txbxContent></wps:txbx>` : "";

  return (
    "<wps:wsp>" +
    cNvSpPr +
    `<wps:spPr bwMode="auto">${spPrParts.join("")}</wps:spPr>` +
    styleXml +
    txbxXml +
    stringifyBodyPr(opts.bodyProperties) +
    "</wps:wsp>"
  );
}

function stringifyNonVisualShapeProperties(opts: NonVisualShapePropertiesOptions): string {
  let xml = "";
  // wps:cNvPr — id/name/descr/title (XSD CT_NonVisualDrawingProps)
  if (opts.id !== undefined || opts.name !== undefined) {
    const attrs: string[] = [];
    if (opts.id !== undefined) attrs.push(`id="${opts.id}"`);
    if (opts.name !== undefined) attrs.push(`name="${escapeXml(opts.name)}"`);
    if (opts.description !== undefined) attrs.push(`descr="${escapeXml(opts.description)}"`);
    if (opts.title !== undefined) attrs.push(`title="${escapeXml(opts.title)}"`);
    xml += `<wps:cNvPr ${attrs.join(" ")}/>`;
  }
  // CT_WordprocessingShape choice: wps:cNvSpPr (text box/autoshape) or
  // wps:cNvCnPr (connector). Connectors carry empty cNvCnPr.
  if (opts.connector) {
    xml += "<wps:cNvCnPr/>";
  } else if (opts.textBox !== undefined) {
    xml += `<wps:cNvSpPr txBox="${opts.textBox}"/>`;
  } else {
    xml += "<wps:cNvSpPr/>";
  }
  return xml;
}

/** Stringify a single style-matrix reference (a:lnRef/a:fillRef/...). */
function stringifyStyleRef(name: string, ref: StyleMatrixReferenceOptions | undefined): string {
  if (!ref) return "";
  const idx = escapeXml(ref.idx);
  const colorXml = ref.color ? createColorElement(ref.color) : "";
  if (colorXml) return `<${name} idx="${idx}">${colorXml}</${name}>`;
  return `<${name} idx="${idx}"/>`;
}

/** Stringify a wps:style (CT_ShapeStyle): lnRef/fillRef/effectRef/fontRef. */
function stringifyShapeStyle(opts: ShapeStyleOptions): string {
  const inner =
    stringifyStyleRef("a:lnRef", opts.lineReference) +
    stringifyStyleRef("a:fillRef", opts.fillReference) +
    stringifyStyleRef("a:effectRef", opts.effectReference) +
    stringifyStyleRef("a:fontRef", opts.fontReference);
  return inner ? `<wps:style>${inner}</wps:style>` : "";
}

function stringifyBodyPr(opts?: BodyPropertiesOptions): string {
  // Delegate to the shared createBodyProperties so attributes + EG_TextAutofit
  // (noAutofit/normAutofit/spAutoFit) + prstTxWarp/3D all round-trip. The old
  // inline copy dropped noAutoFit/spAutoFit and most CT_TextBodyProperties attrs.
  return createBodyProperties(opts ?? {});
}

// ── WPG group (pure string, no class instances) ──

function stringifyWpgGroup(
  opts: {
    children: readonly GroupChildMediaData[];
    transformation: MediaDataTransformation;
    childOffset?: ChildOffset;
    childExtent?: ChildExtent;
    fill?: FillOptions;
    effects?: EffectListOptions;
    groupShapeLocks?: GroupShapeLocksOptions | null;
  },
  ctx: BodyContext,
): string {
  const transform = opts.transformation;
  const grpSpPrParts: string[] = [];
  grpSpPrParts.push(stringifyGroupTransform2D(transform, opts.childOffset, opts.childExtent));
  if (opts.fill) grpSpPrParts.push(fillDesc.stringify(opts.fill, ctx) ?? "");
  if (opts.effects) grpSpPrParts.push(effectListDesc.stringify(opts.effects, NOOP_CTX) ?? "");

  // Children — wps shapes, nested wpg groups, or pic elements
  const childXml = opts.children.map((child) => stringifyGroupChild(child, ctx)).join("");

  return (
    "<wpg:wgp>" +
    stringifyCnvGrpSpPr(opts.groupShapeLocks) +
    `<wpg:grpSpPr>${grpSpPrParts.join("")}</wpg:grpSpPr>` +
    childXml +
    "</wpg:wgp>"
  );
}

/**
 * Stringify one group child: a wps shape, a nested wpg group, or a picture.
 * Shared by the top-level wpg:wgp and nested wpg:grpSp.
 */
function stringifyGroupChild(child: GroupChildMediaData, ctx: BodyContext): string {
  if (child.type === "wps") {
    const wpsData = child as WpsMediaData & { outline?: OutlineOptions; fill?: FillOptions };
    return stringifyWpsShape(
      {
        ...wpsData.data,
        outline: wpsData.outline ?? wpsData.data.outline,
        fill: wpsData.fill ?? wpsData.data.fill,
        transformation: wpsData.transformation,
      },
      ctx,
    );
  }
  if (child.type === "wpg") {
    return stringifyNestedGroup(child as WpgMediaData, ctx);
  }
  // pic child (MediaData) — fill/outline ride on the group-child extension
  // (WpgCommonMediaData) so a grouped picture's spPr round-trips verbatim.
  const picData = child as MediaData & { outline?: OutlineOptions; fill?: FillOptions };
  const picParts: string[] = [];
  picParts.push(stringifyNvPicPr({}, picData.nonVisualProperties));
  const groupBlipParts: string[] = [];
  const useLocalDpiExt = buildUseLocalDpiExt(picData.useLocalDpi);
  groupBlipParts.push(
    useLocalDpiExt
      ? `<a:blip r:embed="{${escapeXml(
          picData.fileName,
        )}}"><a:extLst>${useLocalDpiExt}</a:extLst></a:blip>`
      : `<a:blip r:embed="{${escapeXml(picData.fileName)}}"/>`,
  );
  const groupSrcRectXml = buildSrcRectXml(picData.sourceRectangle);
  if (groupSrcRectXml) groupBlipParts.push(groupSrcRectXml);
  groupBlipParts.push("<a:stretch><a:fillRect/></a:stretch>");
  picParts.push(`<pic:blipFill>${groupBlipParts.join("")}</pic:blipFill>`);
  picParts.push(stringifyShapeProps(picData.transformation, picData.outline, picData.fill));
  return `<pic:pic xmlns:pic="${PIC_URI}">${picParts.join("")}</pic:pic>`;
}

/**
 * Stringify a nested wpg:grpSp (CT_WordprocessingGroup) group child. Same
 * structure as the top-level group, wrapped in wpg:grpSp with a cNvPr id/name.
 */
function stringifyNestedGroup(grp: WpgMediaData, ctx: BodyContext): string {
  const grpSpPrParts: string[] = [];
  grpSpPrParts.push(
    stringifyGroupTransform2D(grp.transformation, grp.childOffset, grp.childExtent),
  );
  if (grp.fill) grpSpPrParts.push(fillDesc.stringify(grp.fill, ctx) ?? "");
  if (grp.effects) grpSpPrParts.push(effectListDesc.stringify(grp.effects, NOOP_CTX) ?? "");
  return (
    "<wpg:grpSp>" +
    '<wpg:cNvPr id="0" name=""/>' +
    stringifyCnvGrpSpPr(grp.groupShapeLocks) +
    `<wpg:grpSpPr>${grpSpPrParts.join("")}</wpg:grpSpPr>` +
    grp.children.map((c) => stringifyGroupChild(c, ctx)).join("") +
    "</wpg:grpSp>"
  );
}

// ── Graphic data content ──

function stringifyGraphicDataContent(
  mediaData: ExtendedMediaData,
  opts: DrawingDescriptorOptions,
  hlIds: HyperlinkIds,
  ctx: BodyContext,
): string {
  const { outline, fill, effects, blipEffects, tile } = opts;
  const transform = mediaData.transformation;

  if (mediaData.type === "chart") {
    const md = mediaData as ChartMediaData;
    return (
      `<a:graphicData uri="${CHART_URI}">` +
      `<c:chart xmlns:c="${CHART_URI}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{chart:${md.chartKey}}"/>` +
      `</a:graphicData>`
    );
  }

  if (mediaData.type === "smartart") {
    const md = mediaData as SmartArtMediaData;
    return (
      `<a:graphicData uri="${DGM_URI}">` +
      `<dgm:relIds xmlns:dgm="${DGM_URI}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:dm="{smartart:${md.smartArtKey}}" r:lo="{smartart-lo:${md.smartArtKey}}" r:qs="{smartart-qs:${md.smartArtKey}}" r:cs="{smartart-cs:${md.smartArtKey}}"/>` +
      `</a:graphicData>`
    );
  }

  if (mediaData.type === "wps") {
    const md = mediaData as WpsMediaData;
    const wpsXml = stringifyWpsShape(
      {
        ...md.data,
        outline,
        fill,
        transformation: transform,
      },
      ctx,
    );
    return `<a:graphicData uri="${WPS_URI}">${wpsXml}</a:graphicData>`;
  }

  if (mediaData.type === "wpg") {
    const md = mediaData as WpgMediaData;
    const wpgXml = stringifyWpgGroup(
      {
        children: md.children,
        transformation: transform,
        childOffset: md.childOffset,
        childExtent: md.childExtent,
        fill: md.fill,
        effects: md.effects,
        groupShapeLocks: md.groupShapeLocks,
      },
      ctx,
    );
    return `<a:graphicData uri="${WPG_URI}">${wpgXml}</a:graphicData>`;
  }

  // Default: image (pic:pic)
  const md = mediaData as MediaData;
  return (
    `<a:graphicData uri="${PIC_URI}">` +
    `<pic:pic xmlns:pic="${PIC_URI}">` +
    stringifyNvPicPr(hlIds, md.nonVisualProperties) +
    stringifyBlipFill(md, blipEffects, tile) +
    stringifyShapeProps(transform, outline, fill, effects) +
    `</pic:pic></a:graphicData>`
  );
}

// ── Position helpers (for anchor) ──

function stringifyPositionH(opts: HorizontalPositionOptions): string {
  const rel = opts.relative ?? HorizontalPositionRelativeFrom.PAGE;
  const child = opts.align
    ? `<wp:align>${opts.align}</wp:align>`
    : opts.offset !== undefined
      ? `<wp:posOffset>${convertToEmu(opts.offset)}</wp:posOffset>`
      : "<wp:align>left</wp:align>";
  return `<wp:positionH relativeFrom="${rel}">${child}</wp:positionH>`;
}

function stringifyPositionV(opts: VerticalPositionOptions): string {
  const rel = opts.relative ?? VerticalPositionRelativeFrom.PAGE;
  const child = opts.align
    ? `<wp:align>${opts.align}</wp:align>`
    : opts.offset !== undefined
      ? `<wp:posOffset>${convertToEmu(opts.offset)}</wp:posOffset>`
      : "<wp:align>top</wp:align>";
  return `<wp:positionV relativeFrom="${rel}">${child}</wp:positionV>`;
}

// ── Text wrapping string builders ──

function wrapPolygonStr(cx: number, cy: number): string {
  return (
    `<wp:wrapPolygon edited="0">` +
    `<wp:start x="0" y="0"/>` +
    `<wp:lineTo x="0" y="${-cy}"/>` +
    `<wp:lineTo x="${cx}" y="${-cy}"/>` +
    `<wp:lineTo x="${cx}" y="0"/>` +
    `<wp:lineTo x="0" y="0"/>` +
    `</wp:wrapPolygon>`
  );
}

function wrapSquareStr(textWrapping: TextWrapping, margins?: Margins): string {
  const side = textWrapping.side ?? TextWrappingSide.BOTH_SIDES;
  const m = margins ?? {};
  const a = [
    `wrapText="${side}"`,
    ...(m.top != null ? [`distT="${convertToEmu(m.top)}"`] : []),
    ...(m.bottom != null ? [`distB="${convertToEmu(m.bottom)}"`] : []),
    ...(m.left != null ? [`distL="${convertToEmu(m.left)}"`] : []),
    ...(m.right != null ? [`distR="${convertToEmu(m.right)}"`] : []),
  ].join(" ");
  return `<wp:wrapSquare ${a}/>`;
}

function wrapTightStr(
  textWrapping: TextWrapping,
  margins: Margins,
  cx: number,
  cy: number,
): string {
  const side = textWrapping.side ?? TextWrappingSide.BOTH_SIDES;
  const a = [`wrapText="${side}"`];
  if (margins.left != null) a.push(`distL="${convertToEmu(margins.left)}"`);
  if (margins.right != null) a.push(`distR="${convertToEmu(margins.right)}"`);
  return `<wp:wrapTight ${a.join(" ")}>${wrapPolygonStr(cx, cy)}</wp:wrapTight>`;
}

function wrapThroughStr(
  textWrapping: TextWrapping,
  margins: Margins,
  cx: number,
  cy: number,
): string {
  const side = textWrapping.side ?? TextWrappingSide.BOTH_SIDES;
  const a = [`wrapText="${side}"`];
  if (margins.left != null) a.push(`distL="${convertToEmu(margins.left)}"`);
  if (margins.right != null) a.push(`distR="${convertToEmu(margins.right)}"`);
  return `<wp:wrapThrough ${a.join(" ")}>${wrapPolygonStr(cx, cy)}</wp:wrapThrough>`;
}

function wrapTopAndBottomStr(margins?: Margins): string {
  const m = margins ?? {};
  const a = [
    ...(m.top != null ? [`distT="${convertToEmu(m.top)}"`] : []),
    ...(m.bottom != null ? [`distB="${convertToEmu(m.bottom)}"`] : []),
  ].join(" ");
  return a ? `<wp:wrapTopAndBottom ${a}/>` : "<wp:wrapTopAndBottom/>";
}

// ── Inline wrapper ──

/** Render wp:cNvGraphicFramePr. Undefined → authoring default (noChangeAspect=1);
 *  `{}` → empty element; otherwise the given lock flags. */
function stringifyCnvGraphicFramePr(locks?: GraphicFrameLocksOptions | null): string {
  const resolved = locks ?? { noChangeAspect: true };
  const attrParts: string[] = [];
  if (resolved.noGrp) attrParts.push('noGrp="1"');
  if (resolved.noDrilldown) attrParts.push('noDrilldown="1"');
  if (resolved.noSelect) attrParts.push('noSelect="1"');
  if (resolved.noChangeAspect) attrParts.push('noChangeAspect="1"');
  if (resolved.noMove) attrParts.push('noMove="1"');
  if (resolved.noResize) attrParts.push('noResize="1"');
  if (attrParts.length === 0) return "<wp:cNvGraphicFramePr/>";
  const attrStr = " " + attrParts.join(" ");
  return `<wp:cNvGraphicFramePr><a:graphicFrameLocks${attrStr} xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></wp:cNvGraphicFramePr>`;
}

/**
 * Stringify wpg:cNvGrpSpPr (CT_NonVisualGroupShapeDrawingProperties): contains
 * an optional a:grpSpLocks (CT_GroupLocking). When no locks are present (the
 * Word default for groups) the element stays empty — groups do NOT inject a
 * default lock, unlike wp:cNvGraphicFramePr.
 */
function stringifyCnvGrpSpPr(locks?: GroupShapeLocksOptions | null): string {
  if (!locks) return "<wpg:cNvGrpSpPr/>";
  const attrParts: string[] = [];
  if (locks.noGrp) attrParts.push('noGrp="1"');
  if (locks.noUngrp) attrParts.push('noUngrp="1"');
  if (locks.noSelect) attrParts.push('noSelect="1"');
  if (locks.noRot) attrParts.push('noRot="1"');
  if (locks.noChangeAspect) attrParts.push('noChangeAspect="1"');
  if (locks.noMove) attrParts.push('noMove="1"');
  if (locks.noResize) attrParts.push('noResize="1"');
  if (attrParts.length === 0) return "<wpg:cNvGrpSpPr/>";
  const attrStr = " " + attrParts.join(" ");
  return `<wpg:cNvGrpSpPr><a:grpSpLocks${attrStr} xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></wpg:cNvGrpSpPr>`;
}

function stringifyInline(
  opts: DrawingDescriptorOptions,
  hlIds: HyperlinkIds,
  ctx: BodyContext,
): string {
  const { mediaData, effects, docProperties } = opts;
  const cx = mediaData.transformation.emus.x;
  const cy = mediaData.transformation.emus.y;

  // Prefer the verbatim source effectExtent (round-trip); fall back to
  // computing it from the shape's effects on the generation path.
  const effectExtent = mediaData.transformation.effectExtent ?? calculateEffectExtent(effects);
  const graphicDataXml = stringifyGraphicDataContent(mediaData, opts, hlIds, ctx);

  return (
    `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:effectExtent l="${effectExtent.l}" t="${effectExtent.t}" r="${effectExtent.r}" b="${effectExtent.b}"/>` +
    stringifyDocPr(docProperties, hlIds) +
    stringifyCnvGraphicFramePr(opts.graphicFrameLocks) +
    `<a:graphic ${GRAPHIC_NS}>${graphicDataXml}</a:graphic>` +
    `</wp:inline></w:drawing>`
  );
}

// ── Anchor (floating) wrapper ──

function stringifyAnchor(
  opts: DrawingDescriptorOptions,
  hlIds: HyperlinkIds,
  ctx: BodyContext,
): string {
  const { mediaData, floating: rawFloating, docProperties } = opts;
  const cx = mediaData.transformation.emus.x;
  const cy = mediaData.transformation.emus.y;

  const floating: Required<Floating> = {
    allowOverlap: true,
    behindDocument: false,
    horizontalPosition: {},
    layoutInCell: true,
    lockAnchor: false,
    verticalPosition: {},
    zIndex: mediaData.transformation.emus.y,
    margins: {},
    wrap: { type: TextWrappingType.NONE },
    ...rawFloating,
  };

  const attrParts = [
    `distT="${convertToEmu(floating.margins?.top ?? 0)}"`,
    `distB="${convertToEmu(floating.margins?.bottom ?? 0)}"`,
    `distL="${convertToEmu(floating.margins?.left ?? 0)}"`,
    `distR="${convertToEmu(floating.margins?.right ?? 0)}"`,
    'simplePos="0"',
    `allowOverlap="${floating.allowOverlap ? 1 : 0}"`,
    `behindDoc="${floating.behindDocument ? 1 : 0}"`,
    `locked="${floating.lockAnchor ? 1 : 0}"`,
    `layoutInCell="${floating.layoutInCell ? 1 : 0}"`,
    `relativeHeight="${floating.zIndex}"`,
  ];

  // Wrap
  let wrapXml: string;
  const rawWrap = rawFloating?.wrap;
  if (rawWrap?.type === TextWrappingType.SQUARE) {
    wrapXml = wrapSquareStr(rawWrap, floating.margins);
  } else if (rawWrap?.type === TextWrappingType.TIGHT) {
    wrapXml = wrapTightStr(rawWrap, floating.margins, cx, cy);
  } else if (rawWrap?.type === TextWrappingType.THROUGH) {
    wrapXml = wrapThroughStr(rawWrap, floating.margins, cx, cy);
  } else if (rawWrap?.type === TextWrappingType.TOP_AND_BOTTOM) {
    wrapXml = wrapTopAndBottomStr(floating.margins);
  } else {
    wrapXml = "<wp:wrapNone/>";
  }

  const graphicDataXml = stringifyGraphicDataContent(mediaData, opts, hlIds, ctx);

  // Prefer the verbatim source effectExtent (round-trip); default to zero.
  const ee = mediaData.transformation.effectExtent;
  const effectExtentXml = ee
    ? `<wp:effectExtent l="${ee.l}" t="${ee.t}" r="${ee.r}" b="${ee.b}"/>`
    : '<wp:effectExtent l="0" t="0" r="0" b="0"/>';

  return (
    `<w:drawing><wp:anchor ${attrParts.join(" ")}>` +
    '<wp:simplePos x="0" y="0"/>' +
    stringifyPositionH(floating.horizontalPosition) +
    stringifyPositionV(floating.verticalPosition) +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    effectExtentXml +
    wrapXml +
    stringifyDocPr(docProperties, hlIds) +
    stringifyCnvGraphicFramePr(opts.graphicFrameLocks) +
    `<a:graphic ${GRAPHIC_NS}>${graphicDataXml}</a:graphic>` +
    `</wp:anchor></w:drawing>`
  );
}

// ── Descriptor ──

/**
 * Drawing descriptor for DOCX `<w:drawing>` elements.
 *
 * Eliminates the Drawing/Inline/Anchor/Graphic/GraphicData/Pic XmlComponent
 * class chain. Inline images, charts, and smartarts produce XML via pure
 * string concatenation — zero XmlComponent instances.
 *
 * @example
 * ```typescript
 * const xml = drawingDesc.stringify({ mediaData, docProperties: opts.altText, floating: opts.floating }, ctx);
 * ```
 */
export const drawingDesc: CustomDescriptor<DrawingDescriptorOptions, BodyContext> = {
  kind: "custom",

  stringify(opts, ctx) {
    // Register hyperlink relationships
    const hlIds = registerHyperlinks(opts.docProperties?.hyperlink, ctx);

    if (opts.floating) {
      return stringifyAnchor(opts, hlIds, ctx);
    }
    return stringifyInline(opts, hlIds, ctx);
  },

  parse(el, ctx) {
    const result = parseDrawingRun(el, ctx as DocxReadContext);
    return (result ?? {}) as unknown as DrawingDescriptorOptions;
  },
};
