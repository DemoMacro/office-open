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

import { type ReproducibleScope } from "@office-open/core";
import { convertToEmu, uniqueNumericIdCreator } from "@office-open/core";
import {
  NOOP_WRITE_CONTEXT,
  type CustomDescriptor,
  type ReadContext,
} from "@office-open/core/descriptor";
import type {
  BlackWhiteMode,
  BlipEffectsOptions,
  EffectListOptions,
  FillOptions,
  OutlineOptions,
  CustomGeometryOptions,
  PresetGeometryOptions,
  OfficeArtExtensionOptions,
  Scene3DOptions,
  Shape3DOptions,
  ShapeType,
  SourceRectangleOptions,
  TileOptions,
} from "@office-open/core/drawing";
import {
  calculateEffectExtent,
  connectorLockingDesc,
  createColorElement,
  groupShapePropertiesDesc,
  pictureLockingDesc,
  shapeLockingDesc,
  shapePropertiesDesc,
  stringifyBlipEffects,
  stringifyBackgroundFormatting,
  stringifyEndpointConnection,
  stringifyNonVisualContentPartProperties,
  stringifyNonVisualDrawingProperties,
  stringifyWholeE2oFormatting,
} from "@office-open/core/drawing";
import { RELATIONSHIP_TYPES, TargetModeType } from "@office-open/core/opc";
import { escapeXml } from "@office-open/xml";
import { stringifyParagraphInline } from "@parts/inline";
import type { ParagraphOptions } from "@parts/paragraph/paragraph";
import { stringifyUnsupportedXmlElementData } from "@parts/paragraph/run/unsupported-drawing-run";
import type {
  ChartMediaData,
  ChartExMediaData,
  ContentPartMediaData,
  ExtendedMediaData,
  GroupChildMediaData,
  LinkedPictureMediaData,
  MediaData,
  MediaTransformation,
  SmartArtMediaData,
  GroupMediaData,
  ShapeMediaData,
  CanvasMediaData,
} from "@shared/media";
import type { NonVisualPropertiesOptions } from "@shared/media/data";
import type { SectionChild } from "@shared/section";

import type { BodyContext, DocxReadContext } from "../../context";
import type { DocumentNamespaceDialect } from "../document/document-attributes";
import type { DocPropertiesOptions, HyperlinkOptions } from "./doc-properties/doc-properties";
import { parseDrawingRun, type DrawingChild } from "./drawing-parse";
import type { Floating, HorizontalPositionOptions, VerticalPositionOptions } from "./floating";
import { HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom } from "./floating";
// wpg/wps types only
import {
  createBodyProperties,
  type BodyPropertiesOptions,
} from "./inline/graphic/graphic-data/wps/body-properties";
import type { NonVisualShapePropertiesOptions } from "./inline/graphic/graphic-data/wps/non-visual-shape-properties";
import type {
  ShapeStyleOptions,
  ShapeStyleReferenceOptions,
  ShapeCoreOptions,
} from "./inline/graphic/graphic-data/wps/wps-shape";
import { TextWrappingSide, TextWrappingType } from "./text-wrap";
import type { TextWrapping, WrapPolygon } from "./text-wrap";

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
  /**
   * The source carried a bare `<a:graphicFrameLocks/>` with no attributes.
   * Round-trip marker only — re-emits the empty element (element presence is
   * part of the source's fidelity) instead of dropping it.
   */
  emptyLocks?: boolean;
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
  /**
   * The source carried a bare `<a:grpSpLocks/>` with no attributes.
   * Round-trip marker only — re-emits the empty element.
   */
  emptyLocks?: boolean;
}

export interface DrawingDescriptorOptions {
  /** Media data (image, chart, smartart, wps, wpg) */
  mediaData: ExtendedMediaData;
  /** Non-visual document properties (name, description, hyperlinks) */
  docProperties?: DocPropertiesOptions;
  /** Whether the drawing wrapper carried wp:docPr; round-trip only. */
  docPropertiesPresent?: boolean;
  /** Floating/anchored positioning (omit for inline) */
  floating?: Floating;
  /** Shape outline */
  outline?: OutlineOptions;
  /** Shape fill */
  fill?: FillOptions;
  /** Shape effects (shadow, glow, etc.) */
  effects?: EffectListOptions;
  /** Preset picture geometry (pic:spPr/a:prstGeom); rect is the authoring default. */
  geometry?: ShapeType | PresetGeometryOptions;
  /** Custom picture geometry (pic:spPr/a:custGeom); wins over geometry. */
  customGeometry?: CustomGeometryOptions;
  /** 3D scene (pic:spPr/a:scene3d) — camera and lighting. */
  scene3d?: Scene3DOptions;
  /** 3D shape properties (pic:spPr/a:sp3d). */
  shape3d?: Shape3DOptions;
  /** Image blip effects (brightness, contrast, etc.) */
  blipEffects?: BlipEffectsOptions;
  /** Raw inner XML of a:blip/a:extLst for non-SVG/non-useLocalDpi extensions (round-trip). */
  blipExt?: string;
  /** Typed a:extLst/a:ext children of pic:spPr (round-trip). */
  shapeExtensions?: { uri: string; content?: string }[];
  /** Image tile fill mode */
  tile?: TileOptions;
  /** Graphic frame locks (wp:cNvGraphicFramePr). `{}` → empty element; omit → authoring default. */
  graphicFrameLocks?: GraphicFrameLocksOptions | null;
  /** Word 2010 wrapper ids; round-trip only. */
  extensionIds?: DrawingExtensionIds;
  /** OOXML namespace dialect; round-trip only. */
  dialect?: DocumentNamespaceDialect;
}

export interface DrawingExtensionIds {
  anchorId?: string;
  editId?: string;
}

// ── ID generation ──

let _docPropsIdGen = uniqueNumericIdCreator();

/** Reset the doc properties ID generator (for testing). */
export const resetDrawingIdGen = (): void => {
  _docPropsIdGen = uniqueNumericIdCreator();
};

// ── Constants ──

const WPS_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const WPG_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
const WPC_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";
const CHART_EX_GRAPHIC_URI = "http://schemas.microsoft.com/office/drawing/2014/chartex";
const CHART_EX_ROOT_NS = "http://schemas.microsoft.com/office/drawing/2014/chartex";
const HYPERLINK_REL = RELATIONSHIP_TYPES.hyperlink;
const IMAGE_REL = RELATIONSHIP_TYPES.image;
const TEXT_BOX_REL = RELATIONSHIP_TYPES.txbxMs;

function drawingExtensionAttrs(ids: DrawingExtensionIds | undefined): string[] {
  const attrs: string[] = [];
  if (ids?.anchorId) attrs.push(`wp14:anchorId="${ids.anchorId}"`);
  if (ids?.editId) attrs.push(`wp14:editId="${ids.editId}"`);
  return attrs;
}
// Blip extension URIs (a:extLst under a:blip).
const SVG_BLIP_EXT_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
const USE_LOCAL_DPI_EXT_URI = "{28A0092B-C50C-407E-A947-70E740481C1C}";
const A14_NS = "http://schemas.microsoft.com/office/drawing/2010/main";

function drawingmlUri(dialect: DocumentNamespaceDialect | undefined, path: string): string {
  return dialect === "strict"
    ? `http://purl.oclc.org/ooxml/drawingml/${path}`
    : `http://schemas.openxmlformats.org/drawingml/2006/${path}`;
}

function relationshipNs(dialect: DocumentNamespaceDialect | undefined): string {
  return dialect === "strict"
    ? "http://purl.oclc.org/ooxml/officeDocument/relationships"
    : "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
}

/**
 * Build the `a14:useLocalDpi` blip extension. Returns "" when the hint is
 * absent (undefined) — Word's default. val="0" (useLocalDpi=false) is the
 * common Word emission; val="1" only when explicitly set.
 */
function buildUseLocalDpiExt(useLocalDpi?: boolean, uri?: string): string {
  if (useLocalDpi === undefined) return "";
  return `<a:ext uri="${escapeXml(uri ?? USE_LOCAL_DPI_EXT_URI)}"><a14:useLocalDpi xmlns:a14="${A14_NS}" val="${
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
  const restoreHyperlink = (
    url: string,
    relationshipId: string | undefined,
    targetMode: "External" | undefined,
  ) => {
    if (relationshipId && /^rId\d+$/.test(relationshipId)) {
      if (
        !ctx.viewWrapper.relationships.hasExactRelationship(
          relationshipId,
          HYPERLINK_REL,
          url,
          targetMode,
        )
      ) {
        ctx.viewWrapper.relationships.addRelationship(
          relationshipId,
          HYPERLINK_REL,
          url,
          targetMode,
        );
      }
      return relationshipId;
    }
    const sourceRel = (ctx.fileData._options?.passthroughRelationships ?? []).find(
      (rel) =>
        rel.source === "word/document.xml" &&
        rel.relationshipType.split("/").pop() === "hyperlink" &&
        rel.target === url,
    );
    const existingId = ctx.viewWrapper.relationships.idOf(HYPERLINK_REL, url);
    if (existingId) return existingId;
    if (sourceRel) return sourceRel.rId;
    return `rId${ctx.viewWrapper.relationships.add(HYPERLINK_REL, url, targetMode)}`;
  };
  if (hyperlink.click) {
    result.clickId = restoreHyperlink(
      hyperlink.click,
      hyperlink.clickRelationshipId,
      hyperlink.clickTargetMode,
    );
  }
  if (hyperlink.hover) {
    result.hoverId = restoreHyperlink(
      hyperlink.hover,
      hyperlink.hoverRelationshipId,
      hyperlink.hoverTargetMode,
    );
  }
  return result;
}

function buildHyperlinkChildren(
  ids: HyperlinkIds,
  hyperlink: HyperlinkOptions | undefined,
  dialect?: DocumentNamespaceDialect,
): string {
  const parts: string[] = [];
  const aNs = `xmlns:a="${drawingmlUri(dialect, "main")}"`;
  if (ids.clickId) {
    parts.push(
      `<a:hlinkClick r:id="${ids.clickId}"${hyperlinkAttributes(hyperlink, "click")} ${aNs}/>`,
    );
  }
  if (ids.hoverId) {
    parts.push(
      `<a:hlinkHover r:id="${ids.hoverId}"${hyperlinkAttributes(hyperlink, "hover")} ${aNs}/>`,
    );
  }
  return parts.join("");
}

function hyperlinkAttributes(
  hyperlink: HyperlinkOptions | undefined,
  kind: "click" | "hover",
): string {
  const tooltip = kind === "click" ? hyperlink?.clickTooltip : hyperlink?.hoverTooltip;
  const targetFrame = kind === "click" ? hyperlink?.clickTargetFrame : hyperlink?.hoverTargetFrame;
  return `${targetFrame === undefined ? "" : ` tgtFrame="${escapeXml(targetFrame)}"`}${tooltip === undefined ? "" : ` tooltip="${escapeXml(tooltip)}"`}`;
}

// ── DocPr ──

function stringifyDocPr(
  opts: DocPropertiesOptions | undefined,
  hlIds: HyperlinkIds,
  hyperlink: HyperlinkOptions | undefined,
  reproducible?: ReproducibleScope,
  dialect?: DocumentNamespaceDialect,
): string {
  const id = opts?.id ?? reproducible?.nextDrawingId() ?? _docPropsIdGen();
  return stringifyNonVisualDrawingProperties(
    "wp:docPr",
    id,
    opts,
    "",
    buildHyperlinkChildren(hlIds, hyperlink, dialect),
  );
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

function buildFillRectXml(fillRect: SourceRectangleOptions | false | undefined): string {
  if (fillRect === false) return "";
  if (!fillRect) return "<a:fillRect/>";
  const attrs: string[] = [];
  if (fillRect.left !== undefined) attrs.push(`l="${fillRect.left}"`);
  if (fillRect.top !== undefined) attrs.push(`t="${fillRect.top}"`);
  if (fillRect.right !== undefined) attrs.push(`r="${fillRect.right}"`);
  if (fillRect.bottom !== undefined) attrs.push(`b="${fillRect.bottom}"`);
  return attrs.length ? `<a:fillRect ${attrs.join(" ")}/>` : "<a:fillRect/>";
}

function stringifyBlipFill(
  mediaData: MediaData | LinkedPictureMediaData,
  blipEffects?: BlipEffectsOptions,
  tile?: TileOptions,
  ctx?: BodyContext,
  blipExt?: string,
): string {
  const fileName =
    mediaData.type === "svg" && "fallback" in mediaData
      ? mediaData.fallback.fileName
      : "fileName" in mediaData
        ? mediaData.fileName
        : undefined;

  const parts: string[] = [];

  // a:blip — cstate omitted unless set; Word's default is "none", so emitting
  // it unconditionally inflates round-trip output that originally had none.
  // A linked-only picture has no embedded copy: r:link alone.
  const blipAttrs: string[] = [];
  if (fileName !== undefined) blipAttrs.push(`r:embed="{${escapeXml(fileName)}}"`);
  if (mediaData.compression !== undefined)
    blipAttrs.push(`cstate="${escapeXml(mediaData.compression)}"`);
  // External linked source (r:link) — a direct External image relationship of
  // the owning part, the same channel docPr hyperlinks use.
  if (mediaData.sourceUrl !== undefined && ctx) {
    const linkId = ctx.viewWrapper.relationships.add(
      IMAGE_REL,
      mediaData.sourceUrl,
      TargetModeType.EXTERNAL,
    );
    blipAttrs.push(`r:link="rId${linkId}"`);
  }

  // Blip extension list: useLocalDpi (rendering hint) + SVG blip reference.
  // Both live in a single shared a:extLst; emitted only when at least one ext
  // is present so blips without extensions stay self-closing.
  let extLstXml: string;
  if (blipExt !== undefined) {
    extLstXml = blipExt ? `<a:extLst>${blipExt}</a:extLst>` : "";
  } else {
    const extParts: string[] = [];
    const useLocalDpiExt = buildUseLocalDpiExt(mediaData.useLocalDpi, mediaData.useLocalDpiUri);
    if (useLocalDpiExt) extParts.push(useLocalDpiExt);
    if (mediaData.type === "svg") {
      extParts.push(
        `<a:ext uri="${SVG_BLIP_EXT_URI}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="{${escapeXml(
          mediaData.fileName,
        )}}"/></a:ext>`,
      );
    }
    extLstXml = extParts.length > 0 ? `<a:extLst>${extParts.join("")}</a:extLst>` : "";
  }

  // Blip effects
  const blipEffectsXml = blipEffects ? stringifyBlipEffects(blipEffects, NOOP_WRITE_CONTEXT) : "";

  // CT_Blip orders the effect choice before the trailing extLst
  const blipContent = blipEffectsXml + extLstXml;
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
  } else if (mediaData.stretch !== false) {
    parts.push(`<a:stretch>${buildFillRectXml(mediaData.fillRectangle)}</a:stretch>`);
  }

  const fillAttrs: string[] = [];
  if (mediaData.dpi !== undefined) fillAttrs.push(`dpi="${mediaData.dpi}"`);
  if (mediaData.rotWithShape !== undefined) {
    fillAttrs.push(`rotWithShape="${mediaData.rotWithShape ? 1 : 0}"`);
  }
  const fillAttr = fillAttrs.length ? ` ${fillAttrs.join(" ")}` : "";
  return `<pic:blipFill${fillAttr}>${parts.join("")}</pic:blipFill>`;
}

// ── Shape Properties (pic:spPr) ──

function stringifyShapeProps(
  transform: MediaTransformation,
  outline?: OutlineOptions,
  fill?: FillOptions,
  effects?: EffectListOptions,
  scene3d?: Scene3DOptions,
  shape3d?: Shape3DOptions,
  blackWhiteMode?: BlackWhiteMode,
  geometry?: ShapeType | PresetGeometryOptions,
  customGeometry?: CustomGeometryOptions,
  extensions?: { uri: string; content?: string }[],
): string {
  const spPr = shapePropertiesDesc.stringify(
    {
      x: convertToEmu(transform.offset?.x ?? 0),
      y: convertToEmu(transform.offset?.y ?? 0),
      width: transform.shapeExtent?.width ?? convertToEmu(transform.width),
      height: transform.shapeExtent?.height ?? convertToEmu(transform.height),
      flipHorizontal: transform.flipHorizontal,
      flipVertical: transform.flipVertical,
      rotation: transform.rotation,
      geometry: customGeometry ? undefined : (geometry ?? "rect"),
      customGeometry,
      fill,
      outline,
      effects,
      scene3d,
      shape3d,
      extensions,
    },
    NOOP_WRITE_CONTEXT,
  );
  const bwAttr = blackWhiteMode ? ` bwMode="${blackWhiteMode}"` : "";
  return `<pic:spPr${bwAttr}>${spPr ?? ""}</pic:spPr>`;
}

// ── Non-visual picture properties (pic:nvPicPr) ──

function stringifyNvPicPr(
  cNvPr: NonVisualPropertiesOptions | undefined,
  ctx: BodyContext,
  dialect?: DocumentNamespaceDialect,
): string {
  const id = cNvPr?.id ?? 0;
  const hyperlinkIds = registerHyperlinks(cNvPr?.hyperlink, ctx);
  const cNvPrXml = stringifyNonVisualDrawingProperties(
    "pic:cNvPr",
    id,
    cNvPr,
    "",
    buildHyperlinkChildren(hyperlinkIds, cNvPr?.hyperlink, dialect),
  );
  // preferRelativeResize defaults to true; only an explicit false is written.
  const cNvPicPrAttr = cNvPr?.preferRelativeResize === false ? ' preferRelativeResize="0"' : "";
  // Picture locks are tri-state: null → source had none (bare pic:cNvPicPr),
  // omitted → authoring default (Word locks the aspect of fresh pictures).
  const locks = cNvPr?.pictureLocks;
  const picLocksXml =
    locks === null
      ? ""
      : (pictureLockingDesc.stringify(locks ?? { noChangeAspect: true }, NOOP_WRITE_CONTEXT) ?? "");
  return `<pic:nvPicPr>${cNvPrXml}<pic:cNvPicPr${cNvPicPrAttr}>${picLocksXml}</pic:cNvPicPr></pic:nvPicPr>`;
}

// ── WPS shape (pure string, no class instances) ──

/** WpsShape options for stringification (extends ShapeCoreOptions with transformation). */
interface WpsStringifyOptions extends ShapeCoreOptions {
  transformation: MediaTransformation;
}

const SECTION_CHILD_KEYS: ReadonlySet<string> = new Set([
  "table",
  "toc",
  "textbox",
  "sdt",
  "altChunk",
  "subDoc",
  "customXml",
  "bookmarkStart",
  "bookmarkEnd",
  "rawXml",
]);

export function stringifyWpsTextBoxChild(
  child: ShapeCoreOptions["children"][number],
  ctx: BodyContext,
): string {
  if (typeof child === "string") return stringifyParagraphInline(child, ctx);
  if ("paragraph" in child) return ctx.stringifyChild(child);
  if (Object.keys(child).some((key) => SECTION_CHILD_KEYS.has(key))) {
    return ctx.stringifyChild(child as SectionChild);
  }
  return stringifyParagraphInline(child as ParagraphOptions, ctx);
}

function stringifyWpsShape(opts: WpsStringifyOptions, ctx: BodyContext): string {
  const transform = opts.transformation;
  const shapeProperties = {
    ...(opts.transformPresent === false
      ? {}
      : {
          x: convertToEmu(opts.shapeOffset?.x ?? transform.offset?.x ?? 0),
          y: convertToEmu(opts.shapeOffset?.y ?? transform.offset?.y ?? 0),
          width: transform.shapeExtent?.width ?? convertToEmu(transform.width),
          height: transform.shapeExtent?.height ?? convertToEmu(transform.height),
          flipHorizontal: transform.flipHorizontal,
          flipVertical: transform.flipVertical,
          rotation: transform.rotation,
        }),
    transformPresent: opts.transformPresent === true ? true : undefined,
    customGeometry: opts.customGeometry,
    // WPS shapes always carry geometry — default to rect when none specified.
    geometry: opts.customGeometry ? undefined : (opts.geometry ?? "rect"),
    fill: opts.fill,
    outline: opts.outline,
    effectDag: opts.effectDag,
    effects: opts.effects,
    scene3d: opts.scene3d,
    shape3d: opts.shape3d,
    extensions: opts.extensions,
  };
  const spPrContent =
    opts.shapePropertiesPresent === false
      ? ""
      : (shapePropertiesDesc.stringify(shapeProperties, ctx) ?? "");
  const spPrXml =
    opts.shapePropertiesPresent === false
      ? ""
      : `<wps:spPr${opts.blackWhiteMode ? ` bwMode="${opts.blackWhiteMode}"` : ""}>${spPrContent}</wps:spPr>`;

  // Non-visual shape properties — default txBox="1"
  const nonVisualXml = opts.nonVisualProperties
    ? stringifyNonVisualShapeProperties(opts.nonVisualProperties, ctx)
    : '<wps:cNvSpPr txBox="1"/>';
  const wrappedXml = opts.wrappedNonVisualProperties
    ? `<wps:nvSpPr>${stringifyNonVisualShapeProperties(opts.wrappedNonVisualProperties, ctx)}</wps:nvSpPr>`
    : "";
  const cNvSpPr =
    opts.nonVisualPropertiesPresent === false
      ? wrappedXml
      : (opts.nonVisualPropertiesWrapped
          ? `<wps:nvSpPr>${nonVisualXml}</wps:nvSpPr>`
          : nonVisualXml) + wrappedXml;

  // CT_TxbxContent contains the full Word block-level group, not only
  // paragraphs: SDTs, tables, custom XML, and unknown raw children all survive.
  const childXml =
    opts.children?.map((child) => stringifyWpsTextBoxChild(child, ctx)).join("") ?? "";

  // Shape style (wps:style) — theme references, emitted after spPr (XSD order)
  const styleXml = opts.style ? stringifyShapeStyle(opts.style) : "";
  // wps:txbx — inline block content or a relationship to a passthrough
  // w14:txbx part. The relationship belongs to the current owner part, not
  // necessarily document.xml (the same shape can live in a header/footer).
  let txbxXml = childXml
    ? `<wps:txbx${opts.textBoxId !== undefined ? ` id="${opts.textBoxId}"` : ""}><w:txbxContent>${childXml}</w:txbxContent></wps:txbx>`
    : "";
  if (!childXml && opts.textBoxPart) {
    const target = textBoxRelationshipTarget(opts.textBoxPart.path, ctx);
    const existingId = ctx.viewWrapper.relationships.idOf(TEXT_BOX_REL, target);
    const relationshipId =
      existingId ?? `rId${ctx.viewWrapper.relationships.add(TEXT_BOX_REL, target)}`;
    txbxXml = `<wps:txbx r:txbx="${relationshipId}" txbxSeq="${opts.textBoxPart.sequence}"/>`;
  }
  // wps:linkedTxbx — XSD choice partner of txbx: the text lives in the linked
  // part, so the shape carries the chain reference instead of inline content.
  const linkedTxbxXml =
    !txbxXml && opts.linkedTextBox
      ? `<wps:linkedTxbx id="${opts.linkedTextBox.id}" seq="${opts.linkedTextBox.sequence}"/>`
      : "";
  // East-Asian vertical flow attribute (default false — emit only when set)
  const neafAttr = opts.normalEastAsianFlow ? ' normalEastAsianFlow="1"' : "";

  return (
    `<wps:wsp${neafAttr}>` +
    cNvSpPr +
    spPrXml +
    styleXml +
    txbxXml +
    linkedTxbxXml +
    (opts.bodyPropertiesPresent === false ? "" : stringifyBodyPr(opts.bodyProperties)) +
    "</wps:wsp>"
  );
}

function textBoxRelationshipTarget(partPath: string, ctx: BodyContext): string {
  if (!partPath.startsWith("word/")) return partPath;
  const ownerPart = ctx.viewWrapper.partName;
  if (!ownerPart?.startsWith("word/")) return partPath.slice("word/".length);
  const ownerSlash = ownerPart.lastIndexOf("/");
  const ownerDirectory = ownerSlash < 0 ? "" : ownerPart.slice(0, ownerSlash + 1);
  return partPath.startsWith(ownerDirectory)
    ? partPath.slice(ownerDirectory.length)
    : partPath.slice("word/".length);
}

function stringifyNonVisualShapeProperties(
  opts: NonVisualShapePropertiesOptions,
  ctx: BodyContext,
): string {
  let xml = "";
  // wps:cNvPr — id/name/descr/title/hidden (XSD CT_NonVisualDrawingProps)
  const hyperlinkIds = registerHyperlinks(opts.hyperlink, ctx);
  const hyperlinkXml = buildHyperlinkChildren(hyperlinkIds, opts.hyperlink);
  if (opts.id !== undefined || opts.name !== undefined || opts.ext || opts.creationId) {
    xml += stringifyNonVisualDrawingProperties("wps:cNvPr", opts.id ?? 0, opts, "", hyperlinkXml);
  } else if (hyperlinkXml) {
    xml += stringifyNonVisualDrawingProperties("wps:cNvPr", 0, opts, "", hyperlinkXml);
  }
  // CT_WordprocessingShape choice: wps:cNvSpPr (text box/autoshape) or
  // wps:cNvCnPr (connector). Each carries its optional locks element.
  if (opts.connector) {
    const connectorLocks = opts.connectorLocking
      ? (connectorLockingDesc.stringify(opts.connectorLocking, NOOP_WRITE_CONTEXT) ?? "")
      : "";
    const connections =
      (opts.startConnection ? stringifyEndpointConnection("stCxn", opts.startConnection) : "") +
      (opts.endConnection ? stringifyEndpointConnection("endCxn", opts.endConnection) : "");
    xml += `<wps:cNvCnPr>${connectorLocks}${connections}</wps:cNvCnPr>`;
  } else {
    const locks = opts.locking
      ? (shapeLockingDesc.stringify(opts.locking, NOOP_WRITE_CONTEXT) ?? "")
      : "";
    const txBox = opts.textBox !== undefined ? ` txBox="${opts.textBox}"` : "";
    xml += locks ? `<wps:cNvSpPr${txBox}>${locks}</wps:cNvSpPr>` : `<wps:cNvSpPr${txBox}/>`;
  }
  return xml;
}

/** Stringify a single style-matrix reference (a:lnRef/a:fillRef/...). */
function stringifyStyleRef(name: string, ref: ShapeStyleReferenceOptions | undefined): string {
  if (!ref) return "";
  const colorXml = ref.color ? createColorElement(ref.color) : "";
  if (colorXml) return `<${name} idx="${ref.index}">${colorXml}</${name}>`;
  return `<${name} idx="${ref.index}"/>`;
}

/** Stringify a:fontRef — @idx is ST_FontCollectionIndex, not a number. */
function stringifyFontRef(ref: ShapeStyleOptions["fontReference"] | undefined): string {
  if (!ref) return "";
  const colorXml = ref.color ? createColorElement(ref.color) : "";
  if (colorXml) return `<a:fontRef idx="${escapeXml(ref.collection)}">${colorXml}</a:fontRef>`;
  return `<a:fontRef idx="${escapeXml(ref.collection)}"/>`;
}

/** Stringify a wps:style (CT_ShapeStyle): lnRef/fillRef/effectRef/fontRef. */
function stringifyShapeStyle(opts: ShapeStyleOptions): string {
  const inner =
    stringifyStyleRef("a:lnRef", opts.lineReference) +
    stringifyStyleRef("a:fillRef", opts.fillReference) +
    stringifyStyleRef("a:effectRef", opts.effectReference) +
    stringifyFontRef(opts.fontReference);
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
    transformation: MediaTransformation;
    childOffsetX?: number;
    childOffsetY?: number;
    childOffsetPresent?: boolean;
    childExtentWidth?: number;
    childExtentHeight?: number;
    childExtentPresent?: boolean;
    groupExtent?: { width: number; height: number };
    extensions?: OfficeArtExtensionOptions[];
    blackWhiteMode?: BlackWhiteMode;
    fill?: FillOptions;
    effects?: EffectListOptions;
    groupShapeLocks?: GroupShapeLocksOptions | null;
    dialect?: DocumentNamespaceDialect;
  },
  ctx: BodyContext,
): string {
  const transform = opts.transformation;
  const grpSpPrContent =
    groupShapePropertiesDesc.stringify(
      {
        x: convertToEmu(transform.offset?.x ?? 0),
        y: convertToEmu(transform.offset?.y ?? 0),
        width: convertToEmu(opts.groupExtent?.width ?? transform.width),
        height: convertToEmu(opts.groupExtent?.height ?? transform.height),
        flipHorizontal: transform.flipHorizontal,
        flipVertical: transform.flipVertical,
        rotation: transform.rotation,
        childOffsetX: opts.childOffsetX,
        childOffsetY: opts.childOffsetY,
        childOffsetPresent: opts.childOffsetPresent,
        childExtentWidth: opts.childExtentWidth,
        childExtentHeight: opts.childExtentHeight,
        childExtentPresent: opts.childExtentPresent,
        fill: opts.fill,
        effects: opts.effects,
      },
      ctx,
    ) ?? "";

  // Children — wps shapes, nested wpg groups, or pic elements
  const childXml = opts.children
    .map((child) => stringifyGroupChild(child, ctx, opts.dialect))
    .join("");
  const extensions = (opts.extensions ?? [])
    .map(
      (extension) => `<a:ext uri="${escapeXml(extension.uri)}">${extension.content ?? ""}</a:ext>`,
    )
    .join("");
  const extensionXml = extensions ? `<a:extLst>${extensions}</a:extLst>` : "";

  return (
    "<wpg:wgp>" +
    stringifyCnvGrpSpPr(opts.groupShapeLocks, opts.dialect) +
    `<wpg:grpSpPr${opts.blackWhiteMode ? ` bwMode="${opts.blackWhiteMode}"` : ""}>${grpSpPrContent}${extensionXml}</wpg:grpSpPr>` +
    childXml +
    "</wpg:wgp>"
  );
}

/** Stringify a WordprocessingCanvas and its shared DrawingML formatting. */
function stringifyWordprocessingCanvas(
  md: CanvasMediaData,
  ctx: BodyContext,
  dialect?: DocumentNamespaceDialect,
): string {
  const background = md.background ? stringifyBackgroundFormatting(md.background, ctx, "wpc") : "";
  const whole = md.whole ? stringifyWholeE2oFormatting(md.whole, ctx, "wpc") : "";
  const extensions = (md.extensions ?? [])
    .map(
      (extension) => `<a:ext uri="${escapeXml(extension.uri)}">${extension.content ?? ""}</a:ext>`,
    )
    .join("");
  const extensionXml = extensions ? `<wpc:extLst>${extensions}</wpc:extLst>` : "";
  const children = md.children.map((child) => stringifyGroupChild(child, ctx, dialect)).join("");
  const wpcXml =
    `<wpc:wpc xmlns:wpc="${WPC_URI}">` +
    background +
    whole +
    children +
    extensionXml +
    "</wpc:wpc>";
  return `<a:graphicData uri="${WPC_URI}">${wpcXml}</a:graphicData>`;
}

/**
 * Stringify one group child: a wps shape, a nested wpg group, or a picture.
 * Shared by the top-level wpg:wgp and nested wpg:grpSp.
 */
function stringifyGroupChild(
  child: GroupChildMediaData,
  ctx: BodyContext,
  dialect?: DocumentNamespaceDialect,
): string {
  if (child.type === "unsupported") {
    return stringifyUnsupportedXmlElementData(child.element, child.relationships, ctx);
  }
  if (child.type === "wps") {
    const wpsData = child as ShapeMediaData & { outline?: OutlineOptions; fill?: FillOptions };
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
    return stringifyNestedGroup(child as GroupMediaData, ctx, dialect);
  }
  if (child.type === "chart") {
    return stringifyGroupGraphicFrame(child as ChartMediaData, dialect);
  }
  if (child.type === "chartEx") {
    return stringifyChartExGraphicFrame(child as ChartExMediaData, dialect);
  }
  if (child.type === "contentPart") {
    return stringifyContentPart("wpg", child as ContentPartMediaData);
  }
  // pic child (MediaData) — fill/outline ride on the group-child extension
  // (GroupCommonMediaData) so a grouped picture's spPr round-trips verbatim.
  const picData = child as MediaData & { outline?: OutlineOptions; fill?: FillOptions };
  const isSvg = picData.type === "svg";
  // a:blip r:embed targets the raster fallback for SVG pictures (what legacy
  // viewers render); the vector SVG lives in the svgBlip extension below.
  const blipTarget = isSvg && "fallback" in picData ? picData.fallback.fileName : picData.fileName;
  const picParts: string[] = [];
  if (picData.nonVisualPropertiesPresent !== false) {
    picParts.push(stringifyNvPicPr(picData.nonVisualProperties, ctx, dialect));
  }
  const groupBlipParts: string[] = [];
  const groupBlipAttrs = [`r:embed="{${escapeXml(blipTarget)}}"`];
  if (picData.compression !== undefined)
    groupBlipAttrs.push(`cstate="${escapeXml(picData.compression)}"`);
  // Verbatim blipExt subsumes useLocalDpi and SVG (see standalone path above).
  let extLst: string;
  if (picData.blipExt !== undefined) {
    extLst = picData.blipExt ? `<a:extLst>${picData.blipExt}</a:extLst>` : "";
  } else {
    const extParts: string[] = [];
    const useLocalDpiExt = buildUseLocalDpiExt(picData.useLocalDpi, picData.useLocalDpiUri);
    if (useLocalDpiExt) extParts.push(useLocalDpiExt);
    if (isSvg) {
      extParts.push(
        `<a:ext uri="${SVG_BLIP_EXT_URI}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="{${escapeXml(
          picData.fileName,
        )}}"/></a:ext>`,
      );
    }
    extLst = extParts.length > 0 ? `<a:extLst>${extParts.join("")}</a:extLst>` : "";
  }
  groupBlipParts.push(
    extLst
      ? `<a:blip ${groupBlipAttrs.join(" ")}>${extLst}</a:blip>`
      : `<a:blip ${groupBlipAttrs.join(" ")}/>`,
  );
  const groupSrcRectXml = buildSrcRectXml(picData.sourceRectangle);
  if (groupSrcRectXml) groupBlipParts.push(groupSrcRectXml);
  if (picData.stretch !== false)
    groupBlipParts.push(`<a:stretch>${buildFillRectXml(picData.fillRectangle)}</a:stretch>`);
  picParts.push(`<pic:blipFill>${groupBlipParts.join("")}</pic:blipFill>`);
  if (picData.shapePropertiesPresent !== false) {
    picParts.push(
      stringifyShapeProps(
        picData.transformation,
        picData.outline,
        picData.fill,
        undefined,
        undefined,
        undefined,
        picData.blackWhiteMode,
        undefined,
        undefined,
        (picData as { shapeExtensions?: { uri: string; content?: string }[] }).shapeExtensions,
      ),
    );
  }
  return `<pic:pic xmlns:pic="${drawingmlUri(dialect, "picture")}">${picParts.join("")}</pic:pic>`;
}

/**
 * Stringify a wpg:graphicFrame group child (CT_GraphicFrame): cNvPr +
 * cNvFrPr + a:xfrm + a:graphic. Charts are the graphic payload Word produces
 * inside groups; the chart part is registered by the group dispatch.
 */
function stringifyGroupGraphicFrame(
  md: ChartMediaData,
  dialect?: DocumentNamespaceDialect,
): string {
  const nvp = md.nonVisualProperties;
  const cNvPrXml = stringifyNonVisualDrawingProperties("wpg:cNvPr", nvp?.id ?? 0, nvp, "Chart");
  const cNvFrPrXml = stringifyCnvFrPr(md.graphicFrameLocks, dialect);
  const xfrmXml = stringifyChildXfrm("wpg", md.transformation);
  return (
    "<wpg:graphicFrame>" +
    cNvPrXml +
    cNvFrPrXml +
    xfrmXml +
    `<a:graphic xmlns:a="${drawingmlUri(dialect, "main")}">` +
    `<a:graphicData uri="${drawingmlUri(dialect, "chart")}">` +
    `<c:chart xmlns:c="${drawingmlUri(dialect, "chart")}" xmlns:r="${relationshipNs(dialect)}" r:id="{chart:${md.chartKey}}"/>` +
    `</a:graphicData>` +
    `</a:graphic>` +
    "</wpg:graphicFrame>"
  );
}

function stringifyChartExGraphicFrame(
  md: ChartExMediaData,
  dialect?: DocumentNamespaceDialect,
): string {
  const nvp = md.nonVisualProperties;
  const cNvPrXml = stringifyNonVisualDrawingProperties("wpg:cNvPr", nvp?.id ?? 0, nvp, "Chart");
  const cNvFrPrXml = stringifyCnvFrPr(md.graphicFrameLocks, dialect);
  const xfrmXml = stringifyChildXfrm("wpg", md.transformation);
  return (
    "<wpg:graphicFrame>" +
    cNvPrXml +
    cNvFrPrXml +
    xfrmXml +
    `<a:graphic xmlns:a="${drawingmlUri(dialect, "main")}">` +
    `<a:graphicData uri="${CHART_EX_GRAPHIC_URI}">` +
    `<cx:chart xmlns:cx="${CHART_EX_ROOT_NS}" xmlns:r="${relationshipNs(dialect)}" r:id="{chartex:${md.chartExKey}}"/>` +
    `</a:graphicData>` +
    `</a:graphic>` +
    "</wpg:graphicFrame>"
  );
}

/** Render wpg:cNvFrPr (CT_NonVisualGraphicFrameProperties — a:graphicFrameLocks). */
function stringifyCnvFrPr(
  locks?: GraphicFrameLocksOptions | null,
  dialect?: DocumentNamespaceDialect,
): string {
  if (!locks) return "<wpg:cNvFrPr/>";
  const attrParts: string[] = [];
  if (locks.noGrp) attrParts.push('noGrp="1"');
  if (locks.noDrilldown) attrParts.push('noDrilldown="1"');
  if (locks.noSelect) attrParts.push('noSelect="1"');
  if (locks.noChangeAspect) attrParts.push('noChangeAspect="1"');
  if (locks.noMove) attrParts.push('noMove="1"');
  if (locks.noResize) attrParts.push('noResize="1"');
  if (attrParts.length === 0) return "<wpg:cNvFrPr/>";
  const attrStr = " " + attrParts.join(" ");
  return `<wpg:cNvFrPr><a:graphicFrameLocks${attrStr} xmlns:a="${drawingmlUri(dialect, "main")}"/></wpg:cNvFrPr>`;
}

/**
 * Build the xfrm (off/ext) of a nested drawing from its transformation. The
 * root tag is host-namespaced (local element of CT_GraphicFrame /
 * CT_WordprocessingContentPart); the off/ext children belong to a:.
 */
function stringifyChildXfrm(prefix: "wp" | "wpg", t: MediaTransformation): string {
  const x = t.offset?.x ?? 0;
  const y = t.offset?.y ?? 0;
  const flipAttrs = (t.flipHorizontal ? ' flipH="1"' : "") + (t.flipVertical ? ' flipV="1"' : "");
  const rotAttr = t.rotation !== undefined ? ` rot="${Math.round(t.rotation * 60000)}"` : "";
  return `<${prefix}:xfrm${flipAttrs}${rotAttr}><a:off x="${x}" y="${y}"/><a:ext cx="${t.width}" cy="${t.height}"/></${prefix}:xfrm>`;
}

/**
 * Stringify a wpg:contentPart child (CT_WordprocessingContentPart) inside a
 * group: nvContentPartPr (cNvPr + cNvContentPartPr) + a:xfrm + `@bwMode`/`@r:id`.
 * A content part never appears at the wp:inline/wp:anchor root (CT_Inline and
 * CT_Anchor take a:graphic only) — the run-level form is w:contentPart.
 */
function stringifyContentPart(prefix: "wpg", md: ContentPartMediaData): string {
  const nvp = md.nonVisualProperties;
  const attrParts = [`r:id="${escapeXml(md.referenceId)}"`];
  if (md.blackWhiteMode) attrParts.push(`bwMode="${escapeXml(md.blackWhiteMode)}"`);
  const nvInner =
    stringifyNonVisualDrawingProperties(`${prefix}:cNvPr`, nvp?.id ?? 0, nvp, "Content Part") +
    stringifyNonVisualContentPartProperties(`${prefix}:cNvContentPartPr`, nvp?.contentPart);
  const nvXml = nvInner ? `<${prefix}:nvContentPartPr>${nvInner}</${prefix}:nvContentPartPr>` : "";
  return (
    `<${prefix}:contentPart ${attrParts.join(" ")}>` +
    nvXml +
    stringifyChildXfrm(prefix, md.transformation) +
    `</${prefix}:contentPart>`
  );
}

/**
 * Stringify a nested wpg:grpSp (CT_WordprocessingGroup) group child. Same
 * structure as the top-level group, wrapped in wpg:grpSp with a cNvPr id/name.
 */
function stringifyNestedGroup(
  grp: GroupMediaData,
  ctx: BodyContext,
  dialect?: DocumentNamespaceDialect,
): string {
  const grpSpPrContent =
    groupShapePropertiesDesc.stringify(
      {
        x: grp.transformation.offset?.x ?? 0,
        y: grp.transformation.offset?.y ?? 0,
        width: grp.transformation.width,
        height: grp.transformation.height,
        flipHorizontal: grp.transformation.flipHorizontal,
        flipVertical: grp.transformation.flipVertical,
        rotation: grp.transformation.rotation,
        childOffsetX: grp.childOffsetX,
        childOffsetY: grp.childOffsetY,
        childOffsetPresent: grp.childOffsetPresent,
        childExtentWidth: grp.childExtentWidth,
        childExtentHeight: grp.childExtentHeight,
        childExtentPresent: grp.childExtentPresent,
        fill: grp.fill,
        effects: grp.effects,
      },
      ctx,
    ) ?? "";
  const extensions = (grp.extensions ?? [])
    .map(
      (extension) => `<a:ext uri="${escapeXml(extension.uri)}">${extension.content ?? ""}</a:ext>`,
    )
    .join("");
  const extensionXml = extensions ? `<a:extLst>${extensions}</a:extLst>` : "";
  return (
    "<wpg:grpSp>" +
    (grp.nonVisualPropertiesPresent === false
      ? ""
      : stringifyNonVisualDrawingProperties(
          "wpg:cNvPr",
          grp.nonVisualProperties?.id ?? 0,
          grp.nonVisualProperties,
          "",
        )) +
    stringifyCnvGrpSpPr(grp.groupShapeLocks, dialect) +
    `<wpg:grpSpPr${grp.blackWhiteMode ? ` bwMode="${grp.blackWhiteMode}"` : ""}>${grpSpPrContent}${extensionXml}</wpg:grpSpPr>` +
    grp.children.map((c) => stringifyGroupChild(c, ctx, dialect)).join("") +
    "</wpg:grpSp>"
  );
}

// ── Graphic data content ──

function stringifyGraphicDataContent(
  mediaData: ExtendedMediaData,
  opts: DrawingDescriptorOptions,
  ctx: BodyContext,
): string {
  const {
    outline,
    fill,
    effects,
    scene3d,
    shape3d,
    blipEffects,
    tile,
    geometry,
    customGeometry,
    blipExt,
    shapeExtensions,
  } = opts;
  const dialect = opts.dialect;
  const transform = mediaData.transformation;

  if (mediaData.type === "chart") {
    const md = mediaData as ChartMediaData;
    return (
      `<a:graphicData uri="${drawingmlUri(dialect, "chart")}">` +
      `<c:chart xmlns:c="${drawingmlUri(dialect, "chart")}" xmlns:r="${relationshipNs(dialect)}" r:id="{chart:${md.chartKey}}"/>` +
      `</a:graphicData>`
    );
  }

  if (mediaData.type === "chartEx") {
    const md = mediaData as ChartExMediaData;
    return (
      `<a:graphicData uri="${CHART_EX_GRAPHIC_URI}">` +
      `<cx:chart xmlns:cx="${CHART_EX_ROOT_NS}" xmlns:r="${relationshipNs(dialect)}" r:id="{chartex:${md.chartExKey}}"/>` +
      `</a:graphicData>`
    );
  }

  if (mediaData.type === "smartart") {
    const md = mediaData as SmartArtMediaData;
    return (
      `<a:graphicData uri="${drawingmlUri(dialect, "diagram")}">` +
      `<dgm:relIds xmlns:dgm="${drawingmlUri(dialect, "diagram")}" xmlns:r="${relationshipNs(dialect)}" r:dm="{smartart:${md.smartArtKey}}" r:lo="{smartart-lo:${md.smartArtKey}}" r:qs="{smartart-qs:${md.smartArtKey}}" r:cs="{smartart-cs:${md.smartArtKey}}"/>` +
      `</a:graphicData>`
    );
  }

  if (mediaData.type === "wps") {
    const md = mediaData as ShapeMediaData;
    const graphicDataUri = md.graphicDataUri ?? WPS_URI;
    const wpsNamespace =
      graphicDataUri === WPS_URI ? "" : ` xmlns:wps="${escapeXml(graphicDataUri)}"`;
    const wpsXml = stringifyWpsShape(
      {
        ...md.data,
        outline: outline ?? md.data.outline,
        fill: fill ?? md.data.fill,
        transformation: transform,
      },
      ctx,
    );
    return `<a:graphicData uri="${escapeXml(graphicDataUri)}"${wpsNamespace}>${wpsXml}</a:graphicData>`;
  }

  if (mediaData.type === "wpg") {
    const md = mediaData as GroupMediaData;
    const wpgXml = stringifyWpgGroup(
      {
        children: md.children,
        transformation: transform,
        groupExtent: md.groupExtent,
        childOffsetX: md.childOffsetX,
        childOffsetY: md.childOffsetY,
        childOffsetPresent: md.childOffsetPresent,
        childExtentWidth: md.childExtentWidth,
        childExtentHeight: md.childExtentHeight,
        childExtentPresent: md.childExtentPresent,
        fill: md.fill,
        extensions: md.extensions,
        blackWhiteMode: md.blackWhiteMode,
        effects: md.effects,
        groupShapeLocks: md.groupShapeLocks,
        dialect,
      },
      ctx,
    );
    return `<a:graphicData uri="${WPG_URI}">${wpgXml}</a:graphicData>`;
  }

  if (mediaData.type === "wpc") {
    return stringifyWordprocessingCanvas(mediaData as CanvasMediaData, ctx, dialect);
  }

  // Default: image (pic:pic)
  const md = mediaData as MediaData | LinkedPictureMediaData;
  return (
    `<a:graphicData uri="${drawingmlUri(dialect, "picture")}">` +
    `<pic:pic xmlns:pic="${drawingmlUri(dialect, "picture")}">` +
    (md.nonVisualPropertiesPresent === false
      ? ""
      : stringifyNvPicPr(md.nonVisualProperties, ctx, dialect)) +
    stringifyBlipFill(md, blipEffects, tile, ctx, blipExt) +
    (md.shapePropertiesPresent === false
      ? ""
      : stringifyShapeProps(
          transform,
          outline,
          fill,
          effects,
          scene3d,
          shape3d,
          md.blackWhiteMode,
          geometry,
          customGeometry,
          shapeExtensions,
        )) +
    `</pic:pic></a:graphicData>`
  );
}

// ── Position helpers (for anchor) ──

function wrapPercentagePosition(
  tag: "wp14:pctPosHOffset" | "wp14:pctPosVOffset",
  positionTag: "wp:positionH" | "wp:positionV",
  relative: string,
  percentOffset: number,
  fallbackOffset: HorizontalPositionOptions["offset"],
): string {
  const choice = `<${positionTag} relativeFrom="${relative}"><${tag}>${Math.round(percentOffset * 1000)}</${tag}></${positionTag}>`;
  if (fallbackOffset === undefined) return choice;
  return (
    '<mc:AlternateContent><mc:Choice Requires="wp14">' +
    choice +
    `</mc:Choice><mc:Fallback><${positionTag} relativeFrom="${relative}"><wp:posOffset>${convertToEmu(fallbackOffset)}</wp:posOffset></${positionTag}></mc:Fallback></mc:AlternateContent>`
  );
}

function stringifyPositionH(opts: HorizontalPositionOptions): string {
  const rel = opts.relative ?? HorizontalPositionRelativeFrom.PAGE;
  if (opts.percentOffset !== undefined) {
    return wrapPercentagePosition(
      "wp14:pctPosHOffset",
      "wp:positionH",
      rel,
      opts.percentOffset,
      opts.offset,
    );
  }
  const child = opts.align
    ? `<wp:align>${opts.align}</wp:align>`
    : opts.offset !== undefined
      ? `<wp:posOffset>${convertToEmu(opts.offset)}</wp:posOffset>`
      : "<wp:align>left</wp:align>";
  return `<wp:positionH relativeFrom="${rel}">${child}</wp:positionH>`;
}

function stringifyPositionV(opts: VerticalPositionOptions): string {
  const rel = opts.relative ?? VerticalPositionRelativeFrom.PAGE;
  if (opts.percentOffset !== undefined) {
    return wrapPercentagePosition(
      "wp14:pctPosVOffset",
      "wp:positionV",
      rel,
      opts.percentOffset,
      opts.offset,
    );
  }
  const child = opts.align
    ? `<wp:align>${opts.align}</wp:align>`
    : opts.offset !== undefined
      ? `<wp:posOffset>${convertToEmu(opts.offset)}</wp:posOffset>`
      : "<wp:align>top</wp:align>";
  return `<wp:positionV relativeFrom="${rel}">${child}</wp:positionV>`;
}

// ── Text wrapping string builders ──

function wrapPolygonStr(cx: number, cy: number, polygon?: WrapPolygon): string {
  // Preserve the source contour verbatim when round-tripped.
  if (polygon?.points.length) {
    if (polygon.points.length < 3) {
      throw new Error("wp:wrapPolygon requires one wp:start and at least two wp:lineTo points");
    }
    // Emit `edited` only when the source had it — keeps the polygon byte-faithful on round-trip.
    const editedAttr = polygon.edited !== undefined ? ` edited="${polygon.edited ? 1 : 0}"` : "";
    const [start, ...rest] = polygon.points;
    // length guard above guarantees `start` exists
    const startStr = `<wp:start x="${start!.x}" y="${start!.y}"/>`;
    const lineToStr = rest.map((p) => `<wp:lineTo x="${p.x}" y="${p.y}"/>`).join("");
    return `<wp:wrapPolygon${editedAttr}>${startStr}${lineToStr}</wp:wrapPolygon>`;
  }
  // Default contour: extent rectangle (origin at top-left, y negated).
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

function wrapSquareStr(textWrapping: TextWrapping): string {
  const side = textWrapping.side ?? TextWrappingSide.BOTH_SIDES;
  const m = textWrapping.margins ?? {};
  const a = [
    `wrapText="${side}"`,
    ...(m.distT != null ? [`distT="${convertToEmu(m.distT)}"`] : []),
    ...(m.distB != null ? [`distB="${convertToEmu(m.distB)}"`] : []),
    ...(m.distL != null ? [`distL="${convertToEmu(m.distL)}"`] : []),
    ...(m.distR != null ? [`distR="${convertToEmu(m.distR)}"`] : []),
  ].join(" ");
  const ee = textWrapping.effectExtent;
  const effectExtentXml = ee
    ? `<wp:effectExtent l="${ee.left}" t="${ee.top}" r="${ee.right}" b="${ee.bottom}"/>`
    : "";
  return ee ? `<wp:wrapSquare ${a}>${effectExtentXml}</wp:wrapSquare>` : `<wp:wrapSquare ${a}/>`;
}

function wrapTightStr(textWrapping: TextWrapping, cx: number, cy: number): string {
  const side = textWrapping.side ?? TextWrappingSide.BOTH_SIDES;
  const a = [`wrapText="${side}"`];
  const m = textWrapping.margins ?? {};
  if (m.distL != null) a.push(`distL="${convertToEmu(m.distL)}"`);
  if (m.distR != null) a.push(`distR="${convertToEmu(m.distR)}"`);
  return `<wp:wrapTight ${a.join(" ")}>${wrapPolygonStr(cx, cy, textWrapping.polygon)}</wp:wrapTight>`;
}

function wrapThroughStr(textWrapping: TextWrapping, cx: number, cy: number): string {
  const side = textWrapping.side ?? TextWrappingSide.BOTH_SIDES;
  const a = [`wrapText="${side}"`];
  const m = textWrapping.margins ?? {};
  if (m.distL != null) a.push(`distL="${convertToEmu(m.distL)}"`);
  if (m.distR != null) a.push(`distR="${convertToEmu(m.distR)}"`);
  return `<wp:wrapThrough ${a.join(" ")}>${wrapPolygonStr(cx, cy, textWrapping.polygon)}</wp:wrapThrough>`;
}

function wrapTopAndBottomStr(textWrapping?: TextWrapping): string {
  const m = textWrapping?.margins ?? {};
  const a = [
    ...(m.distT != null ? [`distT="${convertToEmu(m.distT)}"`] : []),
    ...(m.distB != null ? [`distB="${convertToEmu(m.distB)}"`] : []),
  ].join(" ");
  const ee = textWrapping?.effectExtent;
  const effectExtentXml = ee
    ? `<wp:effectExtent l="${ee.left}" t="${ee.top}" r="${ee.right}" b="${ee.bottom}"/>`
    : "";
  return ee
    ? a
      ? `<wp:wrapTopAndBottom ${a}>${effectExtentXml}</wp:wrapTopAndBottom>`
      : `<wp:wrapTopAndBottom>${effectExtentXml}</wp:wrapTopAndBottom>`
    : a
      ? `<wp:wrapTopAndBottom ${a}/>`
      : "<wp:wrapTopAndBottom/>";
}

// ── Inline wrapper ──

/** Render wp:cNvGraphicFramePr. Null → source had none, omit the element;
 *  undefined → authoring default (noChangeAspect=1); `{}` → source had the
 *  frame without a locks child; `emptyLocks` → bare `<a:graphicFrameLocks/>`;
 *  otherwise the given lock flags. */
function stringifyCnvGraphicFramePr(
  locks?: GraphicFrameLocksOptions | null,
  dialect?: DocumentNamespaceDialect,
): string {
  if (locks === null) return "";
  const resolved = locks ?? { noChangeAspect: true };
  if (resolved.emptyLocks) {
    return `<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${drawingmlUri(dialect, "main")}"/></wp:cNvGraphicFramePr>`;
  }
  const attrParts: string[] = [];
  if (resolved.noGrp) attrParts.push('noGrp="1"');
  if (resolved.noDrilldown) attrParts.push('noDrilldown="1"');
  if (resolved.noSelect) attrParts.push('noSelect="1"');
  if (resolved.noChangeAspect) attrParts.push('noChangeAspect="1"');
  if (resolved.noMove) attrParts.push('noMove="1"');
  if (resolved.noResize) attrParts.push('noResize="1"');
  if (attrParts.length === 0) return "<wp:cNvGraphicFramePr/>";
  const attrStr = " " + attrParts.join(" ");
  return `<wp:cNvGraphicFramePr><a:graphicFrameLocks${attrStr} xmlns:a="${drawingmlUri(dialect, "main")}"/></wp:cNvGraphicFramePr>`;
}

/**
 * Stringify wpg:cNvGrpSpPr (CT_NonVisualGroupShapeDrawingProperties): contains
 * an optional a:grpSpLocks (CT_GroupLocking). When no locks are present (the
 * Word default for groups) the element stays empty — groups do NOT inject a
 * default lock, unlike wp:cNvGraphicFramePr.
 */
function stringifyCnvGrpSpPr(
  locks?: GroupShapeLocksOptions | null,
  dialect?: DocumentNamespaceDialect,
): string {
  if (!locks) return "<wpg:cNvGrpSpPr/>";
  if (locks.emptyLocks) {
    return `<wpg:cNvGrpSpPr><a:grpSpLocks xmlns:a="${drawingmlUri(dialect, "main")}"/></wpg:cNvGrpSpPr>`;
  }
  const attrParts: string[] = [];
  if (locks.noGrp !== undefined) attrParts.push(`noGrp="${locks.noGrp ? 1 : 0}"`);
  if (locks.noUngrp !== undefined) attrParts.push(`noUngrp="${locks.noUngrp ? 1 : 0}"`);
  if (locks.noSelect !== undefined) attrParts.push(`noSelect="${locks.noSelect ? 1 : 0}"`);
  if (locks.noRot !== undefined) attrParts.push(`noRot="${locks.noRot ? 1 : 0}"`);
  if (locks.noChangeAspect !== undefined)
    attrParts.push(`noChangeAspect="${locks.noChangeAspect ? 1 : 0}"`);
  if (locks.noMove !== undefined) attrParts.push(`noMove="${locks.noMove ? 1 : 0}"`);
  if (locks.noResize !== undefined) attrParts.push(`noResize="${locks.noResize ? 1 : 0}"`);
  if (attrParts.length === 0) return "<wpg:cNvGrpSpPr/>";
  const attrStr = " " + attrParts.join(" ");
  return `<wpg:cNvGrpSpPr><a:grpSpLocks${attrStr} xmlns:a="${drawingmlUri(dialect, "main")}"/></wpg:cNvGrpSpPr>`;
}

function stringifyInline(
  opts: DrawingDescriptorOptions,
  hlIds: HyperlinkIds,
  ctx: BodyContext,
): string {
  const { mediaData, effects, docProperties } = opts;
  const cx = convertToEmu(mediaData.transformation.width);
  const cy = convertToEmu(mediaData.transformation.height);

  // Prefer the verbatim source effectExtent (round-trip); fall back to
  // computing it from the shape's effects on the generation path.
  const calculatedEffectExtent = calculateEffectExtent(effects);
  const effectExtent =
    mediaData.transformation.effectExtentPresent === false
      ? undefined
      : (mediaData.transformation.effectExtent ?? {
          left: calculatedEffectExtent.l,
          top: calculatedEffectExtent.t,
          right: calculatedEffectExtent.r,
          bottom: calculatedEffectExtent.b,
        });
  const distances = mediaData.transformation.inlineDistances ?? {};
  const distanceAttrs =
    mediaData.transformation.inlineDistancePresent === false
      ? []
      : [
          `distT="${convertToEmu(distances.top ?? 0)}"`,
          `distB="${convertToEmu(distances.bottom ?? 0)}"`,
          `distL="${convertToEmu(distances.left ?? 0)}"`,
          `distR="${convertToEmu(distances.right ?? 0)}"`,
        ];
  // CT_Inline's choice is a:graphic — a content part only nests inside a wpg
  // group or canvas, never directly under wp:inline.
  const choiceXml = `<a:graphic xmlns:a="${drawingmlUri(
    opts.dialect,
    "main",
  )}">${stringifyGraphicDataContent(mediaData, opts, ctx)}</a:graphic>`;

  return (
    `<w:drawing><wp:inline ${[...distanceAttrs, ...drawingExtensionAttrs(opts.extensionIds)].join(
      " ",
    )}>` +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    (effectExtent
      ? `<wp:effectExtent l="${effectExtent.left}" t="${effectExtent.top}" r="${effectExtent.right}" b="${effectExtent.bottom}"/>`
      : "") +
    stringifyDocPr(
      docProperties,
      hlIds,
      opts.docProperties?.hyperlink,
      ctx.reproducible,
      opts.dialect,
    ) +
    stringifyCnvGraphicFramePr(opts.graphicFrameLocks, opts.dialect) +
    choiceXml +
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
  const cx = convertToEmu(mediaData.transformation.width);
  const cy = convertToEmu(mediaData.transformation.height);

  const floating: Floating = {
    allowOverlap: true,
    behindDocument: false,
    horizontalPosition: {},
    layoutInCell: true,
    lockAnchor: false,
    verticalPosition: {},
    zIndex: convertToEmu(mediaData.transformation.height),
    margins: {},
    wrap: { type: TextWrappingType.NONE },
    ...rawFloating,
  };

  const attrParts = [
    `distT="${convertToEmu(floating.margins?.top ?? 0)}"`,
    `distB="${convertToEmu(floating.margins?.bottom ?? 0)}"`,
    `distL="${convertToEmu(floating.margins?.left ?? 0)}"`,
    `distR="${convertToEmu(floating.margins?.right ?? 0)}"`,
    `simplePos="${floating.simplePos?.enabled ? 1 : 0}"`,
    `allowOverlap="${floating.allowOverlap ? 1 : 0}"`,
    `behindDoc="${floating.behindDocument ? 1 : 0}"`,
    `locked="${floating.lockAnchor ? 1 : 0}"`,
    `layoutInCell="${floating.layoutInCell ? 1 : 0}"`,
    ...(floating.hidden !== undefined ? [`hidden="${floating.hidden ? 1 : 0}"`] : []),
    ...(floating.zIndex !== null ? [`relativeHeight="${floating.zIndex}"`] : []),
    ...drawingExtensionAttrs(opts.extensionIds),
  ];

  // Wrap
  let wrapXml: string;
  const rawWrap = rawFloating?.wrap;
  if (rawFloating?.wrapPresent === false) {
    wrapXml = "";
  } else if (rawWrap?.type === TextWrappingType.SQUARE) {
    wrapXml = wrapSquareStr(rawWrap);
  } else if (rawWrap?.type === TextWrappingType.TIGHT) {
    wrapXml = wrapTightStr(rawWrap, cx, cy);
  } else if (rawWrap?.type === TextWrappingType.THROUGH) {
    wrapXml = wrapThroughStr(rawWrap, cx, cy);
  } else if (rawWrap?.type === TextWrappingType.TOP_AND_BOTTOM) {
    wrapXml = wrapTopAndBottomStr(rawWrap);
  } else {
    wrapXml = "<wp:wrapNone/>";
  }

  // CT_Anchor's choice is a:graphic — a content part only nests inside a wpg
  // group or canvas, never directly under wp:anchor.
  const choiceXml = `<a:graphic xmlns:a="${drawingmlUri(
    opts.dialect,
    "main",
  )}">${stringifyGraphicDataContent(mediaData, opts, ctx)}</a:graphic>`;

  // Prefer the verbatim source effectExtent (round-trip); default to zero.
  const ee = mediaData.transformation.effectExtent;
  const effectExtentXml =
    mediaData.transformation.effectExtentPresent === false
      ? ""
      : ee
        ? `<wp:effectExtent l="${ee.left}" t="${ee.top}" r="${ee.right}" b="${ee.bottom}"/>`
        : '<wp:effectExtent l="0" t="0" r="0" b="0"/>';

  // wp14:sizeRelH/V trail a:graphic in CT_Anchor (Word 2010+); percent is a
  // whole-number percentage in the API, pctWidth/pctHeight carry 1/1000 %.
  const sizeRelH = rawFloating?.horizontalSize;
  const sizeRelV = rawFloating?.verticalSize;
  const sizeRelXml =
    (sizeRelH
      ? `<wp14:sizeRelH${sizeRelH.relative ? ` relativeFrom="${sizeRelH.relative}"` : ""}>` +
        `<wp14:pctWidth>${Math.round((sizeRelH.percent ?? 0) * 1000)}</wp14:pctWidth></wp14:sizeRelH>`
      : "") +
    (sizeRelV
      ? `<wp14:sizeRelV${sizeRelV.relative ? ` relativeFrom="${sizeRelV.relative}"` : ""}>` +
        `<wp14:pctHeight>${Math.round((sizeRelV.percent ?? 0) * 1000)}</wp14:pctHeight></wp14:sizeRelV>`
      : "");

  return (
    `<w:drawing><wp:anchor ${attrParts.join(" ")}>` +
    (rawFloating?.simplePosPresent === false
      ? ""
      : `<wp:simplePos x="${convertToEmu(floating.simplePos?.x ?? 0)}" y="${convertToEmu(
          floating.simplePos?.y ?? 0,
        )}"/>`) +
    (rawFloating?.horizontalPositionPresent === false
      ? ""
      : stringifyPositionH(floating.horizontalPosition)) +
    (rawFloating?.verticalPositionPresent === false
      ? ""
      : stringifyPositionV(floating.verticalPosition)) +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    effectExtentXml +
    wrapXml +
    (opts.docPropertiesPresent === false
      ? ""
      : stringifyDocPr(
          docProperties,
          hlIds,
          opts.docProperties?.hyperlink,
          ctx.reproducible,
          opts.dialect,
        )) +
    stringifyCnvGraphicFramePr(opts.graphicFrameLocks, opts.dialect) +
    choiceXml +
    sizeRelXml +
    `</wp:anchor></w:drawing>`
  );
}

// ── Descriptor ──

function isDocxReadContext(ctx: ReadContext): ctx is DocxReadContext {
  return "docx" in ctx && "currentPart" in ctx;
}

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
export const drawingDesc: CustomDescriptor<
  DrawingDescriptorOptions,
  BodyContext,
  DrawingChild | undefined
> = {
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
    if (!isDocxReadContext(ctx)) {
      throw new TypeError("Drawing parsing requires a DOCX read context");
    }
    const result = parseDrawingRun(el, ctx);
    return result;
  },
};
