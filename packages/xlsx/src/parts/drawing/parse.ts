/**
 * XLSX Drawing — parse helpers for spreadsheetDrawing anchors.
 *
 * @module
 */

import type { UniversalMeasure } from "@office-open/core";
import type { ReadContext } from "@office-open/core/descriptor";
import {
  connectorLockingDesc,
  parseEndpointConnection,
  groupShapePropertiesDesc,
  parseNonVisualDrawingProperties,
  shapePropertiesDesc,
  textBodyDesc,
} from "@office-open/core/drawingml";
import type {
  ConnectorLockingOptions,
  EndpointConnectionOptions,
  NonVisualDrawingPropertiesOptions,
} from "@office-open/core/drawingml";
import { findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type {
  DrawingAnchorOptions,
  DrawingChartOptions,
  DrawingContentPartOptions,
  ConnectorOptions,
  GroupOptions,
  DrawingPictureOptions,
  ShapeOptions,
  GroupConnectorChildOptions,
  GroupShapeChildOptions,
} from "./types";
import { ANCHOR_TYPES } from "./types";
import type { EditAsType } from "./types";

// ── Parse helpers ──

function readNumChild(el: XmlElement, tag: string): number {
  const child = findChild(el, tag);
  if (!child?.elements?.length) return 0;
  const n = Number(child.elements[0]?.text ?? "");
  return Number.isNaN(n) ? 0 : n;
}

interface Marker {
  col: number;
  colOffset?: number | UniversalMeasure;
  row: number;
  rowOffset?: number | UniversalMeasure;
}

function readMarker(el: XmlElement): Marker {
  return {
    col: readNumChild(el, "col") + 1,
    colOffset: readNumChild(el, "colOff") || undefined,
    row: readNumChild(el, "row") + 1,
    rowOffset: readNumChild(el, "rowOff") || undefined,
  };
}

/** Populate anchor fields (col/row/extent/anchorType/editAs/locks) from an anchor element. */
function readAnchorFields(anchor: XmlElement, name: string, result: DrawingAnchorOptions): void {
  result.col = 1;
  result.row = 1;

  const clientData = findChild(anchor, "clientData");
  if (clientData?.attributes) {
    if (clientData.attributes["fLocksWithSheet"] !== undefined) {
      result.locksWithSheet = clientData.attributes["fLocksWithSheet"] !== "0";
    }
    if (clientData.attributes["fPrintsWithSheet"] !== undefined) {
      result.printsWithSheet = clientData.attributes["fPrintsWithSheet"] !== "0";
    }
  }

  const readExt = (el: XmlElement | undefined): void => {
    if (!el?.attributes) return;
    const cx = Number(el.attributes["cx"]);
    const cy = Number(el.attributes["cy"]);
    if (!Number.isNaN(cx)) result.extentCx = cx;
    if (!Number.isNaN(cy)) result.extentCy = cy;
  };

  if (name === "absoluteAnchor") {
    result.anchorType = ANCHOR_TYPES.absolute;
    const pos = findChild(anchor, "pos");
    if (pos?.attributes) {
      const x = Number(pos.attributes["x"]);
      const y = Number(pos.attributes["y"]);
      if (!Number.isNaN(x)) result.absoluteX = x;
      if (!Number.isNaN(y)) result.absoluteY = y;
    }
    readExt(findChild(anchor, "ext"));
    return;
  }

  const from = findChild(anchor, "from");
  if (from) {
    const m = readMarker(from);
    result.col = m.col;
    result.row = m.row;
    if (m.colOffset !== undefined) result.colOffset = m.colOffset;
    if (m.rowOffset !== undefined) result.rowOffset = m.rowOffset;
  }

  if (name === "oneCellAnchor") {
    result.anchorType = ANCHOR_TYPES.oneCell;
    readExt(findChild(anchor, "ext"));
    return;
  }

  // twoCellAnchor
  result.anchorType = ANCHOR_TYPES.twoCell;
  const to = findChild(anchor, "to");
  if (to) {
    const m = readMarker(to);
    result.toCol = m.col;
    result.toRow = m.row;
    if (m.colOffset !== undefined) result.toColOffset = m.colOffset;
    if (m.rowOffset !== undefined) result.toRowOffset = m.rowOffset;
  }
  const editAs = anchor.attributes?.["editAs"] as EditAsType | undefined;
  if (editAs) result.editAs = editAs;
}

/** Read cNvPr (a:CT_NonVisualDrawingProps) from a non-visual properties child. */
function readCNvPr(
  parent: XmlElement,
  nonVisualTag: string,
): Partial<NonVisualDrawingPropertiesOptions> {
  const nonVisual = findChild(parent, nonVisualTag);
  const cNvPr = nonVisual ? findChild(nonVisual, "cNvPr") : undefined;
  return parseNonVisualDrawingProperties(cNvPr);
}

export function parseImageAnchor(
  anchor: XmlElement,
  pic: XmlElement,
  name: string,
): DrawingPictureOptions {
  const result: DrawingPictureOptions = {
    col: 1,
    row: 1,
    rId: readPicRId(pic) ?? "",
  };

  // Actual image size (applies to all anchor types).
  const ext = readPicExtent(pic);
  if (ext.cx !== undefined) result.extentCx = ext.cx;
  if (ext.cy !== undefined) result.extentCy = ext.cy;

  readAnchorFields(anchor, name, result);
  return result;
}

function readPicRId(pic: XmlElement): string | undefined {
  const blipFill = findChild(pic, "blipFill") ?? pic;
  const blip = findChild(blipFill, "a:blip");
  return blip?.attributes?.["r:embed"] as string | undefined;
}

/** Picture extent from pic/spPr/a:xfrm/a:ext (actual image size in EMU). */
function readPicExtent(pic: XmlElement): { cx?: number; cy?: number } {
  const spPr = findChild(pic, "spPr");
  const xfrm = spPr ? findChild(spPr, "a:xfrm") : undefined;
  const ext = xfrm ? findChild(xfrm, "a:ext") : undefined;
  if (!ext?.attributes) return {};
  const cx = Number(ext.attributes["cx"]);
  const cy = Number(ext.attributes["cy"]);
  return {
    cx: Number.isNaN(cx) ? undefined : cx,
    cy: Number.isNaN(cy) ? undefined : cy,
  };
}

export function parseChartAnchor(
  anchor: XmlElement,
  graphicFrame: XmlElement,
): DrawingChartOptions | undefined {
  const graphicData = findChild(
    findChild(graphicFrame, "a:graphic") ?? graphicFrame,
    "a:graphicData",
  );
  const chartEl = graphicData ? findChild(graphicData, "c:chart") : undefined;
  const rId = chartEl?.attributes?.["r:id"] as string | undefined;
  if (!rId) return undefined;

  const result: DrawingChartOptions = { col: 1, row: 1, rId };
  const from = findChild(anchor, "from");
  if (from) {
    const m = readMarker(from);
    result.col = m.col;
    result.row = m.row;
    if (m.colOffset !== undefined) result.colOffset = m.colOffset;
    if (m.rowOffset !== undefined) result.rowOffset = m.rowOffset;
  }
  const clientData = findChild(anchor, "clientData");
  if (clientData?.attributes) {
    if (clientData.attributes["fLocksWithSheet"] !== undefined) {
      result.locksWithSheet = clientData.attributes["fLocksWithSheet"] !== "0";
    }
    if (clientData.attributes["fPrintsWithSheet"] !== undefined) {
      result.printsWithSheet = clientData.attributes["fPrintsWithSheet"] !== "0";
    }
  }
  return result;
}

export function parseShapeAnchor(
  anchor: XmlElement,
  sp: XmlElement,
  name: string,
  ctx: ReadContext,
): ShapeOptions {
  const result = { col: 1, row: 1, spPr: {} } as ShapeOptions;
  readAnchorFields(anchor, name, result);

  Object.assign(result, readCNvPr(sp, "nvSpPr"));

  const spPr = findChild(sp, "spPr");
  if (spPr) result.spPr = shapePropertiesDesc.parse(spPr, ctx);

  const txBody = findChild(sp, "txBody");
  if (txBody) result.textBody = textBodyDesc.parse(txBody, ctx);

  if (sp.attributes?.["macro"] !== undefined) result.macro = String(sp.attributes["macro"]);
  if (sp.attributes?.["textlink"] !== undefined)
    result.textlink = String(sp.attributes["textlink"]);
  return result;
}

/** Read cNvCxnSpPr children (cxnSpLocks/stCxn/endCxn) onto a connector target. */
function readConnectorNonVisual(
  target: {
    locking?: ConnectorLockingOptions;
    startConnection?: EndpointConnectionOptions;
    endConnection?: EndpointConnectionOptions;
  },
  cxnSp: XmlElement,
  ctx: ReadContext,
): void {
  const nvCxnSpPr = findChild(cxnSp, "nvCxnSpPr");
  if (!nvCxnSpPr) return;
  const cNvCxnSpPr = findChild(nvCxnSpPr, "cNvCxnSpPr");
  if (!cNvCxnSpPr) return;
  const cxnSpLocks = findChild(cNvCxnSpPr, "a:cxnSpLocks");
  if (cxnSpLocks) {
    const locks = connectorLockingDesc.parse(cxnSpLocks, ctx);
    if (locks && Object.keys(locks).length > 0) target.locking = locks;
  }
  const stCxn = findChild(cNvCxnSpPr, "a:stCxn");
  if (stCxn) {
    const conn = parseEndpointConnection(stCxn);
    if (conn) target.startConnection = conn;
  }
  const endCxn = findChild(cNvCxnSpPr, "a:endCxn");
  if (endCxn) {
    const conn = parseEndpointConnection(endCxn);
    if (conn) target.endConnection = conn;
  }
}

export function parseConnectorAnchor(
  anchor: XmlElement,
  cxnSp: XmlElement,
  name: string,
  ctx: ReadContext,
): ConnectorOptions {
  const result = { col: 1, row: 1, spPr: {} } as ConnectorOptions;
  readAnchorFields(anchor, name, result);

  Object.assign(result, readCNvPr(cxnSp, "nvCxnSpPr"));

  const spPr = findChild(cxnSp, "spPr");
  if (spPr) result.spPr = shapePropertiesDesc.parse(spPr, ctx);

  if (cxnSp.attributes?.["macro"] !== undefined) result.macro = String(cxnSp.attributes["macro"]);

  readConnectorNonVisual(result, cxnSp, ctx);
  return result;
}

export function parseGroupAnchor(
  anchor: XmlElement,
  grpSp: XmlElement,
  name: string,
  ctx: ReadContext,
): GroupOptions {
  const result = { col: 1, row: 1, grpSpPr: {} } as GroupOptions;
  readAnchorFields(anchor, name, result);

  Object.assign(result, readCNvPr(grpSp, "nvGrpSpPr"));

  const grpSpPrEl = findChild(grpSp, "grpSpPr");
  if (grpSpPrEl) {
    result.grpSpPr = groupShapePropertiesDesc.parse(grpSpPrEl, ctx);
  }

  const shapes: GroupShapeChildOptions[] = [];
  const childConnectors: GroupConnectorChildOptions[] = [];
  for (const child of grpSp.elements ?? []) {
    if (child.name === "sp") {
      const spPr = findChild(child, "spPr");
      const childShape = {
        spPr: spPr ? shapePropertiesDesc.parse(spPr, ctx) : {},
      } as GroupShapeChildOptions;
      Object.assign(childShape, readCNvPr(child, "nvSpPr"));
      const txBody = findChild(child, "txBody");
      if (txBody) childShape.textBody = textBodyDesc.parse(txBody, ctx);
      if (child.attributes?.["macro"] !== undefined)
        childShape.macro = String(child.attributes["macro"]);
      if (child.attributes?.["textlink"] !== undefined)
        childShape.textlink = String(child.attributes["textlink"]);
      shapes.push(childShape);
    } else if (child.name === "cxnSp") {
      const spPr = findChild(child, "spPr");
      const childConn = {
        spPr: spPr ? shapePropertiesDesc.parse(spPr, ctx) : {},
      } as GroupConnectorChildOptions;
      Object.assign(childConn, readCNvPr(child, "nvCxnSpPr"));
      if (child.attributes?.["macro"] !== undefined)
        childConn.macro = String(child.attributes["macro"]);
      readConnectorNonVisual(childConn, child, ctx);
      childConnectors.push(childConn);
    }
  }
  if (shapes.length > 0) result.shapes = shapes;
  if (childConnectors.length > 0) result.connectors = childConnectors;
  return result;
}

export function parseContentPartAnchor(
  anchor: XmlElement,
  contentPart: XmlElement,
  name: string,
): DrawingContentPartOptions | undefined {
  const rId = contentPart.attributes?.["r:id"] as string | undefined;
  if (!rId) return undefined;
  const result = { col: 1, row: 1, rId } as DrawingContentPartOptions;
  readAnchorFields(anchor, name, result);
  return result;
}
