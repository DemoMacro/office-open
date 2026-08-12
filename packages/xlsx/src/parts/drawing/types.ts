/**
 * XLSX Drawing — anchor object types.
 *
 * Option interfaces for spreadsheetDrawing anchors: images, charts, shapes,
 * connectors, groups, and content parts bound to worksheet cells.
 *
 * @module
 */

import type { UniversalMeasure } from "@office-open/core";
import type {
  ConnectorLockingOptions,
  EndpointConnectionOptions,
  GroupTransform2DOptions,
  ShapePropertiesOptions,
  TextBodyOptions,
} from "@office-open/core/drawingml";

// ── Types (used by compiler) ──

export interface ImageOptions {
  /** 1-based column */
  col: number;
  /** Column offset in EMU (default 0) */
  colOffset?: number | UniversalMeasure;
  /** 1-based row */
  row: number;
  /** Row offset in EMU (default 0) */
  rowOffset?: number | UniversalMeasure;
  /** Relationship ID for the image */
  rId: string;
  /** Lock anchor with sheet (default true) */
  locksWithSheet?: boolean;
  /** Print with sheet (default true) */
  printsWithSheet?: boolean;
}

export interface ChartAnchorOptions {
  /** 1-based column */
  col: number;
  /** Column offset in EMU (default 0) */
  colOffset?: number | UniversalMeasure;
  /** 1-based row */
  row: number;
  /** Row offset in EMU (default 0) */
  rowOffset?: number | UniversalMeasure;
  /** Relationship ID for the chart */
  rId: string;
  /** Lock anchor with sheet (default true) */
  locksWithSheet?: boolean;
  /** Print with sheet (default true) */
  printsWithSheet?: boolean;
}

// ── Descriptor Types ──

/** How a drawing is anchored to the worksheet (xdr:*Anchor element). */
export const ANCHOR_TYPES = {
  twoCell: "twoCell",
  oneCell: "oneCell",
  absolute: "absolute",
} as const;
export type AnchorType = (typeof ANCHOR_TYPES)[keyof typeof ANCHOR_TYPES];

/** editAs behavior for twoCellAnchor (ST_EditAs). */
export const EDIT_AS_TYPES = {
  twoCell: "twoCell",
  oneCell: "oneCell",
  absolute: "absolute",
} as const;
export type EditAsType = (typeof EDIT_AS_TYPES)[keyof typeof EDIT_AS_TYPES];

/** Shared anchor fields for all anchored drawing objects. */
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
  /** Print with sheet (default true) */
  printsWithSheet?: boolean;
}

export interface DrawingPictureOptions extends DrawingAnchorOptions {
  /** Relationship ID for the image */
  rId: string;
}

export interface DrawingChartOptions {
  /** 1-based column */
  col: number;
  /** Column offset in EMU (default 0) */
  colOffset?: number | UniversalMeasure;
  /** 1-based row */
  row: number;
  /** Row offset in EMU (default 0) */
  rowOffset?: number | UniversalMeasure;
  /** Relationship ID for the chart */
  rId: string;
  /** Lock anchor with sheet (default true) */
  locksWithSheet?: boolean;
  /** Print with sheet (default true) */
  printsWithSheet?: boolean;
}

/** Anchored shape (xdr:sp): geometry + optional text body. */
export interface ShapeOptions extends DrawingAnchorOptions {
  /** Shape name (cNvPr name). Defaults to "Shape <id>". */
  name?: string;
  /** Shape properties (a:CT_ShapeProperties). */
  spPr: ShapePropertiesOptions;
  /** Text body (a:CT_TextBody). */
  textBody?: TextBodyOptions;
  /** macro attribute (CT_Shape). */
  macro?: string;
  /** textlink attribute (CT_Shape). */
  textlink?: string;
}

/** Anchored connector (xdr:cxnSp): line/arrow geometry via spPr. */
export interface ConnectorOptions extends DrawingAnchorOptions {
  /** Connector name. Defaults to "Connector <id>". */
  name?: string;
  /** Shape properties (a:CT_ShapeProperties, typically prstGeom="line"). */
  spPr: ShapePropertiesOptions;
  /** macro attribute (CT_Connector). */
  macro?: string;
  /** a:cxnSpLocks — connector locking (inside cNvCxnSpPr). */
  locking?: ConnectorLockingOptions;
  /** a:stCxn — start endpoint glued to a shape connection site. */
  startConnection?: EndpointConnectionOptions;
  /** a:endCxn — end endpoint glued to a shape connection site. */
  endConnection?: EndpointConnectionOptions;
}

/** Shape nested inside a group (no anchor — positioned via spPr.xfrm). */
export interface GroupShapeChildOptions {
  name?: string;
  spPr: ShapePropertiesOptions;
  textBody?: TextBodyOptions;
  macro?: string;
  textlink?: string;
}

/** Connector nested inside a group (no anchor). */
export interface GroupConnectorChildOptions {
  name?: string;
  spPr: ShapePropertiesOptions;
  macro?: string;
  /** a:cxnSpLocks — connector locking (inside cNvCxnSpPr). */
  locking?: ConnectorLockingOptions;
  /** a:stCxn — start endpoint glued to a shape connection site. */
  startConnection?: EndpointConnectionOptions;
  /** a:endCxn — end endpoint glued to a shape connection site. */
  endConnection?: EndpointConnectionOptions;
}

/** Anchored group (xdr:grpSp): group transform + nested shapes/connectors. */
export interface GroupOptions extends DrawingAnchorOptions {
  /** Group name. Defaults to "Group <id>". */
  name?: string;
  /** Group shape properties (a:CT_GroupShapeProperties: group xfrm + fill/ln). */
  grpSpPr: GroupTransform2DOptions;
  /** Nested shapes. */
  shapes?: GroupShapeChildOptions[];
  /** Nested connectors. */
  connectors?: GroupConnectorChildOptions[];
}

/** Anchored external content reference (xdr:contentPart, r:id only). */
export interface DrawingContentPartOptions extends DrawingAnchorOptions {
  /** Relationship ID for the external content. */
  rId: string;
}

export interface DrawingOptions {
  images?: DrawingPictureOptions[];
  charts?: DrawingChartOptions[];
  shapes?: ShapeOptions[];
  connectors?: ConnectorOptions[];
  groups?: GroupOptions[];
  contentParts?: DrawingContentPartOptions[];
}
