/**
 * XLSX Drawing — anchor object types.
 *
 * Option interfaces for spreadsheetDrawing anchors: images, charts, shapes,
 * connectors, groups, and content parts bound to worksheet cells.
 *
 * @module
 */

import type { BaseConnectorOptions, BaseGroupOptions, UniversalMeasure } from "@office-open/core";
import type {
  BlackWhiteMode,
  BlipEffectsOptions,
  GraphicFrameLockingOptions,
  BlipCompression,
  GroupTransform2DOptions,
  NonVisualDrawingPropertiesOptions,
  PictureLockingOptions,
  ShapeLockingOptions,
  ShapePropertiesOptions,
  SourceRectangleOptions,
  TextBodyOptions,
  TextHyperlinkOptions,
} from "@office-open/core/drawing";
import type { DefaultShapeStyleOptions } from "@office-open/core/theme";

// ── Types (used by compiler) ──

// ImageOptions/ChartAnchorOptions were removed: the compiler builds the same
// Drawing*Options types the descriptor consumes, carrying the full anchor set.

// ── Descriptor Types ──

/** How a drawing is anchored to the worksheet (xdr:*Anchor element). */
export const ANCHOR_TYPES = {
  twoCell: "twoCell",
  oneCell: "oneCell",
  absolute: "absolute",
} as const;
/** How a drawing is anchored to the worksheet: "twoCell" from/to cell markers, "oneCell" from-marker plus size, "absolute" absolute EMU offset. */
export type AnchorType = (typeof ANCHOR_TYPES)[keyof typeof ANCHOR_TYPES];

/** editAs behavior for twoCellAnchor (ST_EditAs). */
export const EDIT_AS_TYPES = {
  twoCell: "twoCell",
  oneCell: "oneCell",
  absolute: "absolute",
} as const;
/** How a twoCellAnchor tracks later row/column edits: "twoCell" move and resize, "oneCell" move only, "absolute" neither. */
export type EditAsType = (typeof EDIT_AS_TYPES)[keyof typeof EDIT_AS_TYPES];

/**
 * Shared anchor fields for all anchored drawing objects. 1-based col/row for
 * authoring convenience (the XML marker is 0-based; the descriptor subtracts).
 * Note anchors (NoteAnchorOptions) mirror the XML's 0-based CT_Marker instead.
 */
export interface DrawingAnchorOptions {
  /** 1-based column (from marker) */
  col: number;
  /** Column offset in EMU (default 0) */
  colOffset?: number | UniversalMeasure;
  /** 1-based row (from marker) */
  row: number;
  /** Row offset in EMU (default 0) */
  rowOffset?: number | UniversalMeasure;
  /** To cell column (1-based) for twoCellAnchor. Defaults to col + 1. */
  toCol?: number;
  /** To cell row (1-based) for twoCellAnchor. Defaults to row + 1. */
  toRow?: number;
  /** To cell column offset in EMU. */
  toColOffset?: number | UniversalMeasure;
  /** To cell row offset in EMU. */
  toRowOffset?: number | UniversalMeasure;
  /** Anchor type (default "twoCell"). */
  anchorType?: AnchorType;
  /** editAs for twoCellAnchor (default "oneCell"). */
  editAs?: EditAsType;
  /** Absolute X in EMU (absoluteAnchor). */
  absoluteX?: number | UniversalMeasure;
  /** Absolute Y in EMU (absoluteAnchor). */
  absoluteY?: number | UniversalMeasure;
  /** Anchor extent width in EMU (oneCell/absoluteAnchor ext, default 400000). */
  extentCx?: number | UniversalMeasure;
  /** Anchor extent height in EMU (oneCell/absoluteAnchor ext, default 300000). */
  extentCy?: number | UniversalMeasure;
  /** Lock anchor with sheet (default true) */
  locksWithSheet?: boolean;
  /**
   * Source wrapped the anchor in mc:AlternateContent. `"choice"`: the Choice
   * carries the full anchor and the Fallback is empty (Excel 2010+ a14 form).
   * `"fallback"`: the Choice is an empty `Requires="v"` marker and the
   * Fallback carries the full anchor (legacy VML-compat comment form).
   */
  alternateContent?: "choice" | "fallback";
  /**
   * Source wrapped the anchored object (not the anchor) in
   * mc:AlternateContent/mc:Choice (Excel 2010+ a14 textboxes). The anchor
   * footprint stays unwrapped; only the object is inside the wrapper.
   */
  objectAlternateContent?: boolean;
  /** Print with sheet (default true) */
  printsWithSheet?: boolean;
  /**
   * Document-order position of this anchor inside the drawing part — the
   * z-order when objects overlap. Round-trip only: stringify emits anchors
   * in this order; fresh authoring keeps the per-type bucket order.
   */
  zOrder?: number;
  /** Original cNvPr id (round-trip only; fresh output allocates ids). */
  shapeId?: number;
}

/** Pick the anchor fields defined on `source` (undefined ones stay absent). */
export function pickAnchorOptions<T extends DrawingAnchorOptions>(source: T): DrawingAnchorOptions {
  const picked: DrawingAnchorOptions = { col: source.col, row: source.row };
  const keys = [
    "colOffset",
    "rowOffset",
    "toCol",
    "toRow",
    "toColOffset",
    "toRowOffset",
    "anchorType",
    "editAs",
    "absoluteX",
    "absoluteY",
    "extentCx",
    "extentCy",
    "locksWithSheet",
    "printsWithSheet",
  ] as const;
  // Correlated union-key writes need this cast — TS cannot narrow the write
  // type from a `keys` element alone.
  const optional = picked as unknown as Record<(typeof keys)[number], unknown>;
  for (const key of keys) {
    if (source[key] !== undefined) optional[key] = source[key];
  }
  return picked;
}

export interface DrawingPictureOptions
  extends DrawingAnchorOptions, NonVisualDrawingPropertiesOptions {
  /**
   * Relationship ID for the embedded image (a:blip @r:embed). Empty on a
   * linked-only picture, which carries {@link linkRId} alone.
   */
  rId: string;
  /**
   * Relationship ID for the linked source (a:blip @r:link) — an External
   * image relationship targeting the source URL. Absent when the picture is
   * purely embedded.
   */
  linkRId?: string;
  /**
   * Round-tripped pic/spPr — carries rotation/flip/bwMode/fill that the
   * position-only default emission would drop. When absent, stringify emits
   * the standard xfrm + rect geometry.
   */
  properties?: ShapePropertiesOptions;
  /**
   * Click hyperlink on the object itself (a:hlinkClick inside xdr:cNvPr) —
   * jump to a URL when the object is clicked.
   */
  hyperlink?: TextHyperlinkOptions;
  /** Blip crop (a:srcRect); an empty object round-trips the bare marker. */
  sourceRectangle?: SourceRectangleOptions;
  /** Display resolution hint (blipFill/@dpi), e.g. 300 = 300 DPI. */
  dpi?: number;
  /** Whether the fill rotates with the shape (blipFill/@rotWithShape). */
  rotWithShape?: boolean;
  /** Black/white mode (spPr/@bwMode); absent = attribute omitted. */
  blackWhiteMode?: BlackWhiteMode;
  /** Published-object flag (pic/@fPublished); explicit true/false round-trips. */
  fPublished?: boolean;
  /** Picture locks (cNvPicPr/a:picLocks); absent = empty cNvPicPr. */
  locking?: PictureLockingOptions;
  /**
   * Relative-resize hint (cNvPicPr/@preferRelativeResize). Absent = attribute
   * omitted (defaults true); explicit true/false round-trips the attribute.
   */
  preferRelativeResize?: boolean;
  /** Image adjustment effects carried inside a:blip (a:lum, a:duotone, …). */
  blipEffects?: BlipEffectsOptions;
  /** Compression state (a:blip/@cstate); absent = attribute omitted. */
  compression?: BlipCompression;
  /** Local-DPI display hint (a14:useLocalDpi in the a:blip extension list). */
  useLocalDpi?: boolean;
  /**
   * Verbatim a:blip `a:extLst` inner XML for extensions beyond useLocalDpi
   * (a14 imgProps artistic effects, …). Round-trip only; subsumes
   * `useLocalDpi` when the source list carries both. Named blipExt to
   * avoid colliding with the inherited cNvPr `ext` channel.
   */
  blipExt?: string;
  /**
   * Verbatim cNvPicPr `a:extLst` inner XML for unmodeled extensions
   * (a14:cameraTool, …). Round-trip only — do not hand-author.
   */
  cNvPicPrExt?: string;
}

export interface DrawingChartOptions
  extends DrawingAnchorOptions, NonVisualDrawingPropertiesOptions {
  /** Relationship ID for the chart */
  rId: string;
  /** Graphic-frame transform width (a:xfrm/a:ext/@cx, EMU). */
  frameExtentCx?: number;
  /** Graphic-frame transform height (a:xfrm/a:ext/@cy, EMU). */
  frameExtentCy?: number;
  /** Graphic-frame rotation in degrees (xdr:xfrm/@rot). */
  frameRotation?: number;
  /** Horizontal flip (xdr:xfrm/@flipH). */
  frameFlipHorizontal?: boolean;
  /** Vertical flip (xdr:xfrm/@flipV). */
  frameFlipVertical?: boolean;
  /**
   * Click hyperlink on the object itself (a:hlinkClick inside xdr:cNvPr) —
   * jump to a URL when the object is clicked.
   */
  hyperlink?: TextHyperlinkOptions;
  /** Frame locks (cNvGraphicFramePr/a:graphicFrameLocks); absent = empty. */
  frameLocks?: GraphicFrameLockingOptions;
  /** Macro reference (CT_GraphicFrame/@macro); empty string round-trips. */
  macro?: string;
  /** Published-object flag (graphicFrame/@fPublished); explicit true/false round-trips. */
  fPublished?: boolean;
}

export interface DrawingWebExtensionFallbackOptions {
  /** Relationship ID of the fallback snapshot image. */
  rId: string;
  /** Snapshot picture properties (spPr). */
  properties?: ShapePropertiesOptions;
  /** Black/white mode (spPr/@bwMode); absent = attribute omitted. */
  blackWhiteMode?: BlackWhiteMode;
  /** Compression state (a:blip/@cstate); absent = attribute omitted. */
  compression?: BlipCompression;
  /** Blip crop (a:srcRect); an empty object round-trips the bare marker. */
  sourceRectangle?: SourceRectangleOptions;
  /** Image adjustment effects carried inside a:blip. */
  blipEffects?: BlipEffectsOptions;
  /** Local-DPI display hint (a14:useLocalDpi in the a:blip extension list). */
  useLocalDpi?: boolean;
  /** Verbatim a:blip a:extLst inner XML for unmodeled extensions. */
  blipExt?: string;
  /**
   * Verbatim cNvPicPr `a:extLst` inner XML for unmodeled extensions
   * (a14:cameraTool, …). Round-trip only — do not hand-author.
   */
  cNvPicPrExt?: string;
  /** Picture locks (cNvPicPr/a:picLocks); absent = empty cNvPicPr. */
  locking?: PictureLockingOptions;
  /** Relative-resize hint (cNvPicPr/@preferRelativeResize). */
  preferRelativeResize?: boolean;
  /** Published-object flag (pic/@fPublished). */
  fPublished?: boolean;
}

export interface DrawingWebExtensionOptions
  extends DrawingAnchorOptions, NonVisualDrawingPropertiesOptions {
  /** Relationship ID for the WebExtension part (we:webextensionref/@r:id). */
  rId: string;
  /** Source graphic-data child spelling; `webextensionref` by default. */
  elementName?: "webextension" | "webextensionref";
  /** Snapshot picture used by the mc:Fallback branch. */
  fallback?: DrawingWebExtensionFallbackOptions;
  /** Frame locks (cNvGraphicFramePr/a:graphicFrameLocks); absent = empty. */
  frameLocks?: GraphicFrameLockingOptions;
  /** Macro reference (CT_GraphicFrame/@macro); empty string round-trips. */
  macro?: string;
  /** Published-object flag (graphicFrame/@fPublished). */
  fPublished?: boolean;
}

/**
 * Anchored SmartArt (xdr:graphicFrame whose graphicData carries dgm:relIds).
 * The four diagram parts (data/layout/quickStyle/colors) travel as passthrough
 * parts; these ids wire the frame to them in the drawing rels.
 */
export interface DrawingSmartArtOptions
  extends DrawingAnchorOptions, NonVisualDrawingPropertiesOptions {
  /** Diagram data part (dgm:relIds @r:dm). */
  dataRId: string;
  /** Diagram layout part (dgm:relIds @r:lo). */
  layoutRId: string;
  /** Diagram quick style part (dgm:relIds @r:qs). */
  quickStyleRId: string;
  /** Diagram colors part (dgm:relIds @r:cs). */
  colorsRId: string;
  /** Frame locks (cNvGraphicFramePr/a:graphicFrameLocks); absent = empty. */
  frameLocks?: GraphicFrameLockingOptions;
  /** Macro reference (CT_GraphicFrame/@macro); empty string round-trips. */
  macro?: string;
  /** Published-object flag (graphicFrame/@fPublished); explicit true/false round-trips. */
  fPublished?: boolean;
}

/** Anchored shape (xdr:sp): geometry + optional text body. */
export interface ShapeOptions extends DrawingAnchorOptions, NonVisualDrawingPropertiesOptions {
  /**
   * Click hyperlink on the object itself (a:hlinkClick inside xdr:cNvPr) —
   * jump to a URL when the object is clicked.
   */
  hyperlink?: TextHyperlinkOptions;
  /** Shape locks (cNvSpPr/a:spLocks, CT_ShapeLocking). */
  locking?: ShapeLockingOptions;
  /** Shape properties (a:CT_ShapeProperties). */
  properties: ShapePropertiesOptions;
  /** Text body (a:CT_TextBody). */
  textBody?: TextBodyOptions;
  /** Theme style-matrix references (xdr:style, CT_ShapeStyle). */
  style?: DefaultShapeStyleOptions;
  /** macro attribute (CT_Shape). */
  macro?: string;
  /** textlink attribute (CT_Shape). */
  textlink?: string;
  /** Text-lock flag (CT_Shape `@fLocksText`); explicit value round-trips. */
  fLocksText?: boolean;
  /** Published-object flag (sp/@fPublished); explicit true/false round-trips. */
  fPublished?: boolean;
  /** Text-box flag (cNvSpPr/@txBox) — emitted only when the source had it. */
  textBox?: boolean;
  /** Black/white mode (spPr/@bwMode); absent = attribute omitted. */
  blackWhiteMode?: BlackWhiteMode;
}

/** Anchored connector (xdr:cxnSp): line/arrow geometry via spPr. */
export interface ConnectorOptions extends DrawingAnchorOptions, BaseConnectorOptions {
  /**
   * Click hyperlink on the object itself (a:hlinkClick inside xdr:cNvPr) —
   * jump to a URL when the object is clicked.
   */
  hyperlink?: TextHyperlinkOptions;
  /** Shape properties (a:CT_ShapeProperties, typically prstGeom="line"). */
  properties: ShapePropertiesOptions;
  /** Theme style-matrix references (xdr:style, CT_ShapeStyle). */
  style?: DefaultShapeStyleOptions;
  /** macro attribute (CT_Connector). */
  macro?: string;
  /** Published-object flag (cxnSp/@fPublished); explicit true/false round-trips. */
  fPublished?: boolean;
  /** Black/white mode (spPr/@bwMode); absent = attribute omitted. */
  blackWhiteMode?: BlackWhiteMode;
}

/** Shape nested inside a group (no anchor — positioned via spPr.xfrm). */
export interface GroupShapeChildOptions extends NonVisualDrawingPropertiesOptions {
  /**
   * Click hyperlink on the object itself (a:hlinkClick inside xdr:cNvPr) —
   * jump to a URL when the object is clicked.
   */
  hyperlink?: TextHyperlinkOptions;
  properties: ShapePropertiesOptions;
  textBody?: TextBodyOptions;
  /** Theme style-matrix references (xdr:style, CT_ShapeStyle). */
  style?: DefaultShapeStyleOptions;
  macro?: string;
  textlink?: string;
  /** Shape locks (cNvSpPr/a:spLocks, CT_ShapeLocking). */
  locking?: ShapeLockingOptions;
  /** Text-box flag (cNvSpPr/@txBox) — emitted only when the source had it. */
  textBox?: boolean;
  /** Black/white mode (spPr/@bwMode); absent = attribute omitted. */
  blackWhiteMode?: BlackWhiteMode;
  /** Original cNvPr id (round-trip only; fresh output allocates ids). */
  shapeId?: number;
  /** Published-object flag (sp/@fPublished); explicit true/false round-trips. */
  fPublished?: boolean;
}

/** Connector nested inside a group (no anchor). */
export interface GroupConnectorChildOptions extends BaseConnectorOptions {
  /**
   * Click hyperlink on the object itself (a:hlinkClick inside xdr:cNvPr) —
   * jump to a URL when the object is clicked.
   */
  hyperlink?: TextHyperlinkOptions;
  properties: ShapePropertiesOptions;
  /** Theme style-matrix references (xdr:style, CT_ShapeStyle). */
  style?: DefaultShapeStyleOptions;
  macro?: string;
  /** Black/white mode (spPr/@bwMode); absent = attribute omitted. */
  blackWhiteMode?: BlackWhiteMode;
  /** Original cNvPr id (round-trip only; fresh output allocates ids). */
  shapeId?: number;
  /** Published-object flag (cxnSp/@fPublished); explicit true/false round-trips. */
  fPublished?: boolean;
}

/** Anchored group (xdr:grpSp): group transform + nested shapes/connectors. */
export interface GroupOptions extends DrawingAnchorOptions, BaseGroupOptions {
  /**
   * Click hyperlink on the object itself (a:hlinkClick inside xdr:cNvPr) —
   * jump to a URL when the object is clicked.
   */
  hyperlink?: TextHyperlinkOptions;
  /** Group shape properties (a:CT_GroupShapeProperties: group xfrm + fill/ln). */
  properties: GroupTransform2DOptions;
  /** Black/white mode (grpSpPr/@bwMode); absent = attribute omitted. */
  blackWhiteMode?: BlackWhiteMode;
  /** Nested shapes. */
  shapes?: GroupShapeChildOptions[];
  /** Nested pictures. */
  images?: GroupPictureChildOptions[];
  /** Nested connectors. */
  connectors?: GroupConnectorChildOptions[];
}

/** Picture nested inside a group (no worksheet anchor, but size/id survive). */
export type GroupPictureChildOptions = Omit<DrawingPictureOptions, keyof DrawingAnchorOptions> &
  Pick<DrawingPictureOptions, "extentCx" | "extentCy" | "shapeId">;

/** Anchored external content reference (xdr:contentPart, r:id only). */
export interface DrawingContentPartOptions extends DrawingAnchorOptions {
  /** Relationship ID for the external content. */
  rId: string;
}

export interface DrawingOptions {
  images?: DrawingPictureOptions[];
  charts?: DrawingChartOptions[];
  /** Anchored WebExtension frames (we:webextensionref inside mc:AlternateContent). */
  webExtensions?: DrawingWebExtensionOptions[];
  smartArts?: DrawingSmartArtOptions[];
  shapes?: ShapeOptions[];
  connectors?: ConnectorOptions[];
  groups?: GroupOptions[];
  contentParts?: DrawingContentPartOptions[];
}
