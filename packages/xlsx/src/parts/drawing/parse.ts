/**
 * XLSX Drawing — parse helpers for spreadsheetDrawing anchors.
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import type { UniversalMeasure } from "@office-open/core";
import type { ReadContext } from "@office-open/core/descriptor";
import type { BlackWhiteMode } from "@office-open/core/drawing";
import {
  blipDesc,
  connectorLockingDesc,
  graphicFrameLockingDesc,
  pictureLockingDesc,
  parseEndpointConnection,
  groupShapePropertiesDesc,
  parseNonVisualDrawingProperties,
  readHyperlink,
  shapePropertiesDesc,
  sourceRectangleDesc,
  textBodyDesc,
} from "@office-open/core/drawing";
import type {
  ConnectorLockingOptions,
  EndpointConnectionOptions,
  NonVisualDrawingPropertiesOptions,
  TextHyperlinkOptions,
} from "@office-open/core/drawing";
import { parseShapeStyle } from "@office-open/core/theme";
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

/**
 * SpreadsheetDrawing child lookup tolerant of both namespace forms Office
 * writes: elements in the default namespace (our own output) and elements
 * carrying the xdr: prefix (most external files). The DrawingML (a:) and
 * chart (c:) children inside are always prefixed and go through findChild.
 */
export function findXdr(el: XmlElement, local: string): XmlElement | undefined {
  return findChild(el, local) ?? findChild(el, `xdr:${local}`);
}

function readNumChild(el: XmlElement, tag: string): number {
  const child = findXdr(el, tag);
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

  const clientData = findXdr(anchor, "clientData");
  if (clientData?.attributes) {
    if (clientData.attributes["fLocksWithSheet"] !== undefined) {
      result.locksWithSheet = parseOnOff(clientData.attributes["fLocksWithSheet"]) ?? true;
    }
    if (clientData.attributes["fPrintsWithSheet"] !== undefined) {
      result.printsWithSheet = parseOnOff(clientData.attributes["fPrintsWithSheet"]) ?? true;
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
    const pos = findXdr(anchor, "pos");
    if (pos?.attributes) {
      const x = Number(pos.attributes["x"]);
      const y = Number(pos.attributes["y"]);
      if (!Number.isNaN(x)) result.absoluteX = x;
      if (!Number.isNaN(y)) result.absoluteY = y;
    }
    readExt(findXdr(anchor, "ext"));
    return;
  }

  const from = findXdr(anchor, "from");
  if (from) {
    const m = readMarker(from);
    result.col = m.col;
    result.row = m.row;
    if (m.colOffset !== undefined) result.colOffset = m.colOffset;
    if (m.rowOffset !== undefined) result.rowOffset = m.rowOffset;
  }

  if (name === "oneCellAnchor") {
    result.anchorType = ANCHOR_TYPES.oneCell;
    readExt(findXdr(anchor, "ext"));
    return;
  }

  // twoCellAnchor
  result.anchorType = ANCHOR_TYPES.twoCell;
  const to = findXdr(anchor, "to");
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
  ctx: ReadContext,
): Partial<NonVisualDrawingPropertiesOptions> & { hyperlink?: TextHyperlinkOptions } {
  const nonVisual = findXdr(parent, nonVisualTag);
  const cNvPr = nonVisual ? findXdr(nonVisual, "cNvPr") : undefined;
  const result: Partial<NonVisualDrawingPropertiesOptions> & {
    hyperlink?: TextHyperlinkOptions;
  } = parseNonVisualDrawingProperties(cNvPr);
  const hlinkClick = cNvPr ? findChild(cNvPr, "a:hlinkClick") : undefined;
  if (hlinkClick) result.hyperlink = readHyperlink(hlinkClick, ctx);
  return result;
}

export function parseImageAnchor(
  anchor: XmlElement,
  pic: XmlElement,
  name: string,
  ctx: ReadContext,
): DrawingPictureOptions {
  const refs = readPicRefs(pic);
  const result: DrawingPictureOptions = {
    col: 1,
    row: 1,
    rId: refs.embed ?? "",
    ...(refs.link ? { linkRId: refs.link } : {}),
  };
  Object.assign(result, readCNvPr(pic, "nvPicPr", ctx));

  // preferRelativeResize (defaults true) and the a:blip adjustment effects.
  const blipFill = findXdr(pic, "blipFill");
  const nvPicPr = findXdr(pic, "nvPicPr");
  const cNvPicPr = nvPicPr ? findXdr(nvPicPr, "cNvPicPr") : undefined;
  if (cNvPicPr?.attributes?.["preferRelativeResize"] !== undefined) {
    result.preferRelativeResize =
      parseOnOff(String(cNvPicPr.attributes["preferRelativeResize"])) ?? true;
  }
  if (cNvPicPr) {
    const locks = findChild(cNvPicPr, "a:picLocks");
    if (locks) result.locking = pictureLockingDesc.parse(locks, ctx);
  }
  const blip = blipFill ? findChild(blipFill, "a:blip") : undefined;
  if (blip) {
    const parsed = blipDesc.parse(blip, {} as never);
    if (parsed.blipEffects) result.blipEffects = parsed.blipEffects;
  }
  const srcRect = blipFill ? findChild(blipFill, "a:srcRect") : undefined;
  if (srcRect) result.sourceRectangle = sourceRectangleDesc.parse(srcRect, ctx);

  // Full spPr (rotation/flip/fill) beyond the position-only default; @bwMode
  // is a container attribute the descriptor leaves to the caller.
  const spPrEl = findXdr(pic, "spPr");
  if (spPrEl) {
    result.spPr = shapePropertiesDesc.parse(spPrEl, ctx);
    const bwMode = spPrEl.attributes?.["bwMode"];
    if (bwMode !== undefined) result.blackWhiteMode = bwMode as BlackWhiteMode;
  }

  // Actual image size (applies to all anchor types).
  const ext = readPicExtent(pic);
  if (ext.cx !== undefined) result.extentCx = ext.cx;
  if (ext.cy !== undefined) result.extentCy = ext.cy;

  readAnchorFields(anchor, name, result);
  return result;
}

/** Picture blip references: r:embed (local copy) and/or r:link (external source). */
function readPicRefs(pic: XmlElement): { embed?: string; link?: string } {
  const blipFill = findXdr(pic, "blipFill") ?? pic;
  const attrs = findChild(blipFill, "a:blip")?.attributes;
  return {
    embed: attrs?.["r:embed"] as string | undefined,
    link: attrs?.["r:link"] as string | undefined,
  };
}

/** Picture extent from pic/spPr/a:xfrm/a:ext (actual image size in EMU). */
function readPicExtent(pic: XmlElement): { cx?: number; cy?: number } {
  const spPr = findXdr(pic, "spPr");
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
  name: string,
  ctx: ReadContext,
): DrawingChartOptions | undefined {
  const graphicData = findChild(
    findChild(graphicFrame, "a:graphic") ?? graphicFrame,
    "a:graphicData",
  );
  const chartEl = graphicData ? findChild(graphicData, "c:chart") : undefined;
  const rId = chartEl?.attributes?.["r:id"] as string | undefined;
  if (!rId) return undefined;

  const result = { col: 1, row: 1, rId } as DrawingChartOptions;
  Object.assign(result, readCNvPr(graphicFrame, "nvGraphicFramePr", ctx));
  const nvGraphicFramePr = findXdr(graphicFrame, "nvGraphicFramePr");
  const cNvGraphicFramePr = nvGraphicFramePr
    ? findXdr(nvGraphicFramePr, "cNvGraphicFramePr")
    : undefined;
  if (cNvGraphicFramePr) {
    const locks = findChild(cNvGraphicFramePr, "a:graphicFrameLocks");
    if (locks) result.frameLocks = graphicFrameLockingDesc.parse(locks, ctx);
  }
  if (graphicFrame.attributes?.["macro"] !== undefined)
    result.macro = String(graphicFrame.attributes["macro"]);

  readAnchorFields(anchor, name, result);
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

  Object.assign(result, readCNvPr(sp, "nvSpPr", ctx));
  const nvSpPr = findXdr(sp, "nvSpPr");
  const cNvSpPr = nvSpPr ? findXdr(nvSpPr, "cNvSpPr") : undefined;
  if (cNvSpPr?.attributes?.["txBox"] !== undefined)
    result.textBox = parseOnOff(String(cNvSpPr.attributes["txBox"])) ?? true;

  const spPr = findXdr(sp, "spPr");
  if (spPr) result.spPr = shapePropertiesDesc.parse(spPr, ctx);

  const styleEl = findXdr(sp, "style");
  if (styleEl) {
    const style = parseShapeStyle(styleEl, ctx);
    if (style) result.style = style;
  }

  const txBody = findXdr(sp, "txBody");
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
  const nvCxnSpPr = findXdr(cxnSp, "nvCxnSpPr");
  if (!nvCxnSpPr) return;
  const cNvCxnSpPr = findXdr(nvCxnSpPr, "cNvCxnSpPr");
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

  Object.assign(result, readCNvPr(cxnSp, "nvCxnSpPr", ctx));

  const spPr = findXdr(cxnSp, "spPr");
  if (spPr) result.spPr = shapePropertiesDesc.parse(spPr, ctx);

  if (cxnSp.attributes?.["macro"] !== undefined) result.macro = String(cxnSp.attributes["macro"]);

  readConnectorNonVisual(result, cxnSp, ctx);
  const topConnStyle = findXdr(cxnSp, "style");
  if (topConnStyle) {
    const style = parseShapeStyle(topConnStyle, ctx);
    if (style) result.style = style;
  }
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

  Object.assign(result, readCNvPr(grpSp, "nvGrpSpPr", ctx));

  const grpSpPrEl = findXdr(grpSp, "grpSpPr");
  if (grpSpPrEl) {
    result.grpSpPr = groupShapePropertiesDesc.parse(grpSpPrEl, ctx);
  }

  const shapes: GroupShapeChildOptions[] = [];
  const childConnectors: GroupConnectorChildOptions[] = [];
  for (const child of grpSp.elements ?? []) {
    // Group children appear in the default namespace (our own output) or with
    // the xdr: prefix (external files) — match by local name.
    const local = child.name?.startsWith("xdr:") ? child.name.slice(4) : child.name;
    if (local === "sp") {
      const spPr = findXdr(child, "spPr");
      const childShape = {
        spPr: spPr ? shapePropertiesDesc.parse(spPr, ctx) : {},
      } as GroupShapeChildOptions;
      Object.assign(childShape, readCNvPr(child, "nvSpPr", ctx));
      const childNvSpPr = findXdr(child, "nvSpPr");
      const childCnVSpPr = childNvSpPr ? findXdr(childNvSpPr, "cNvSpPr") : undefined;
      if (childCnVSpPr?.attributes?.["txBox"] !== undefined)
        childShape.textBox = parseOnOff(String(childCnVSpPr.attributes["txBox"])) ?? true;
      const childStyle = findXdr(child, "style");
      if (childStyle) {
        const style = parseShapeStyle(childStyle, ctx);
        if (style) childShape.style = style;
      }
      const txBody = findXdr(child, "txBody");
      if (txBody) childShape.textBody = textBodyDesc.parse(txBody, ctx);
      if (child.attributes?.["macro"] !== undefined)
        childShape.macro = String(child.attributes["macro"]);
      if (child.attributes?.["textlink"] !== undefined)
        childShape.textlink = String(child.attributes["textlink"]);
      shapes.push(childShape);
    } else if (local === "cxnSp") {
      const spPr = findXdr(child, "spPr");
      const childConn = {
        spPr: spPr ? shapePropertiesDesc.parse(spPr, ctx) : {},
      } as GroupConnectorChildOptions;
      Object.assign(childConn, readCNvPr(child, "nvCxnSpPr", ctx));
      if (child.attributes?.["macro"] !== undefined)
        childConn.macro = String(child.attributes["macro"]);
      readConnectorNonVisual(childConn, child, ctx);
      const connStyle = findXdr(child, "style");
      if (connStyle) {
        const style = parseShapeStyle(connStyle, ctx);
        if (style) childConn.style = style;
      }
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
