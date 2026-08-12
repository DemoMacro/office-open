/**
 * XLSX Drawing — stringify helpers for spreadsheetDrawing anchors.
 *
 * @module
 */

import { convertToEmu } from "@office-open/core";
import type { UniversalMeasure } from "@office-open/core";
import type { WriteContext } from "@office-open/core/descriptor";
import {
  connectorLockingDesc,
  groupShapePropertiesDesc,
  shapePropertiesDesc,
  stringifyEndpointConnection,
  stringifyNonVisualDrawingProperties,
  textBodyDesc,
} from "@office-open/core/drawingml";
import type {
  ConnectorLockingOptions,
  EndpointConnectionOptions,
  NonVisualDrawingPropertiesOptions,
  ShapePropertiesOptions,
  TextBodyOptions,
} from "@office-open/core/drawingml";
import { escapeXml } from "@office-open/xml";

import type {
  DrawingAnchorOptions,
  DrawingContentPartOptions,
  ConnectorOptions,
  GroupOptions,
  DrawingPictureOptions,
  DrawingChartOptions,
  ShapeOptions,
} from "./types";
import { ANCHOR_TYPES, EDIT_AS_TYPES } from "./types";

// ── Constants ──

export const XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
export const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
export const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const C_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";

export const DEFAULT_EXTENT_CX = 400000;
export const DEFAULT_EXTENT_CY = 300000;

// ── Stringify helpers ──

/** Marker cell (0-based col/row + EMU offsets). */
function markerXml(
  col: number,
  colOff: number | UniversalMeasure,
  row: number,
  rowOff: number | UniversalMeasure,
): string {
  return `<col>${col - 1}</col><colOff>${convertToEmu(colOff)}</colOff><row>${row - 1}</row><rowOff>${convertToEmu(rowOff)}</rowOff>`;
}

export function clientDataXml(obj: {
  locksWithSheet?: boolean;
  printsWithSheet?: boolean;
}): string {
  const locks = obj.locksWithSheet !== false ? 1 : 0;
  const prints = obj.printsWithSheet !== false ? 1 : 0;
  return `<clientData fLocksWithSheet="${locks}" fPrintsWithSheet="${prints}"/>`;
}

/** Wrap an anchored object in the appropriate xdr:*Anchor element. */
export function wrapAnchor(opts: DrawingAnchorOptions, inner: string): string {
  const anchorType = opts.anchorType ?? ANCHOR_TYPES.twoCell;
  const cx = convertToEmu(opts.extentCx ?? DEFAULT_EXTENT_CX);
  const cy = convertToEmu(opts.extentCy ?? DEFAULT_EXTENT_CY);

  if (anchorType === ANCHOR_TYPES.absolute) {
    const x = convertToEmu(opts.absoluteX ?? 0);
    const y = convertToEmu(opts.absoluteY ?? 0);
    return `<absoluteAnchor><pos x="${x}" y="${y}"/><ext cx="${cx}" cy="${cy}"/>${inner}</absoluteAnchor>`;
  }

  const from = markerXml(opts.col, opts.colOffset ?? 0, opts.row, opts.rowOffset ?? 0);

  if (anchorType === ANCHOR_TYPES.oneCell) {
    return `<oneCellAnchor><from>${from}</from><ext cx="${cx}" cy="${cy}"/>${inner}</oneCellAnchor>`;
  }

  // twoCell
  const editAs = opts.editAs ?? EDIT_AS_TYPES.oneCell;
  const to = markerXml(
    opts.toCol ?? opts.col + 1,
    opts.toColOffset ?? 0,
    opts.toRow ?? opts.row + 1,
    opts.toRowOffset ?? 0,
  );
  return `<twoCellAnchor editAs="${editAs}"><from>${from}</from><to>${to}</to>${inner}</twoCellAnchor>`;
}

function picXml(
  img: DrawingPictureOptions,
  id: number,
  cx: number,
  cy: number,
  ctx: WriteContext,
): string {
  const spPr =
    shapePropertiesDesc.stringify({ x: 0, y: 0, width: cx, height: cy, geometry: "rect" }, ctx) ??
    "";
  return (
    `<pic><nvPicPr>${stringifyNonVisualDrawingProperties("cNvPr", id, img, `Picture ${id}`)}<cNvPicPr preferRelativeResize="1"/></nvPicPr>` +
    `<blipFill><a:blip r:embed="${img.rId}"/><a:stretch><a:fillRect/></a:stretch></blipFill>` +
    `<spPr>${spPr}</spPr></pic>`
  );
}

export function stringifyImage(img: DrawingPictureOptions, id: number, ctx: WriteContext): string {
  const cx = convertToEmu(img.extentCx ?? DEFAULT_EXTENT_CX);
  const cy = convertToEmu(img.extentCy ?? DEFAULT_EXTENT_CY);
  const pic = picXml(img, id, cx, cy, ctx);
  return wrapAnchor(img, `${pic}${clientDataXml(img)}`);
}

export function stringifyChart(chart: DrawingChartOptions, id: number): string {
  // Charts default to twoCellAnchor (existing behavior).
  const from = markerXml(chart.col, chart.colOffset ?? 0, chart.row, chart.rowOffset ?? 0);
  const to = markerXml(chart.col + 9, 0, chart.row + 16, 0);
  const clientData = clientDataXml(chart);
  return (
    `<twoCellAnchor editAs="oneCell"><from>${from}</from><to>${to}</to>` +
    `<graphicFrame><nvGraphicFramePr>${stringifyNonVisualDrawingProperties("cNvPr", id, chart, `Chart ${id}`)}` +
    `<cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></cNvGraphicFramePr></nvGraphicFramePr>` +
    `<xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xfrm>` +
    `<a:graphic><a:graphicData uri="${C_URI}">` +
    `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${R_NS}" r:id="${chart.rId}"/>` +
    `</a:graphicData></a:graphic></graphicFrame>${clientData}</twoCellAnchor>`
  );
}

/** Build the inner xdr:sp content (nvSpPr + spPr + optional txBody). */
function buildShapeContent(
  cNvPr: NonVisualDrawingPropertiesOptions | undefined,
  id: number,
  fallbackName: string,
  spPr: ShapePropertiesOptions,
  textBody: TextBodyOptions | undefined,
  ctx: WriteContext,
  attrs = "",
): string {
  const spPrXml = shapePropertiesDesc.stringify(spPr, ctx) ?? "";
  const txBodyXml = textBody ? `<txBody>${textBodyDesc.stringify(textBody, ctx)}</txBody>` : "";
  return `<sp${attrs}><nvSpPr>${stringifyNonVisualDrawingProperties("cNvPr", id, cNvPr, fallbackName)}<cNvSpPr/></nvSpPr><spPr>${spPrXml}</spPr>${txBodyXml}</sp>`;
}

/** Build the inner xdr:cxnSp content (nvCxnSpPr + spPr). */
function buildConnectorContent(
  cNvPr: NonVisualDrawingPropertiesOptions | undefined,
  id: number,
  fallbackName: string,
  spPr: ShapePropertiesOptions,
  ctx: WriteContext,
  attrs = "",
  connector?: {
    locking?: ConnectorLockingOptions;
    startConnection?: EndpointConnectionOptions;
    endConnection?: EndpointConnectionOptions;
  },
): string {
  const spPrXml = shapePropertiesDesc.stringify(spPr, ctx) ?? "";
  const cNvCxnSpPrInner: string[] = [];
  if (connector?.locking) {
    const locks = connectorLockingDesc.stringify(connector.locking, ctx);
    if (locks) cNvCxnSpPrInner.push(locks);
  }
  if (connector?.startConnection) {
    cNvCxnSpPrInner.push(stringifyEndpointConnection("stCxn", connector.startConnection));
  }
  if (connector?.endConnection) {
    cNvCxnSpPrInner.push(stringifyEndpointConnection("endCxn", connector.endConnection));
  }
  const cNvCxnSpPr = cNvCxnSpPrInner.length
    ? `<cNvCxnSpPr>${cNvCxnSpPrInner.join("")}</cNvCxnSpPr>`
    : "<cNvCxnSpPr/>";
  return `<cxnSp${attrs}><nvCxnSpPr>${stringifyNonVisualDrawingProperties("cNvPr", id, cNvPr, fallbackName)}${cNvCxnSpPr}</nvCxnSpPr><spPr>${spPrXml}</spPr></cxnSp>`;
}

export function stringifyShape(shape: ShapeOptions, id: number, ctx: WriteContext): string {
  const xml = buildShapeContent(
    shape,
    id,
    `Shape ${id}`,
    shape.spPr,
    shape.textBody,
    ctx,
    macroTextlinkAttrs(shape),
  );
  return wrapAnchor(shape, `${xml}${clientDataXml(shape)}`);
}

export function stringifyConnector(conn: ConnectorOptions, id: number, ctx: WriteContext): string {
  const xml = buildConnectorContent(
    conn,
    id,
    `Connector ${id}`,
    conn.spPr,
    ctx,
    macroTextlinkAttrs(conn),
    conn,
  );
  return wrapAnchor(conn, `${xml}${clientDataXml(conn)}`);
}

export function stringifyContentPart(cp: DrawingContentPartOptions): string {
  return wrapAnchor(cp, `<contentPart r:id="${cp.rId}"/>${clientDataXml(cp)}`);
}

/** Build xdr:grpSp content and return the next available cNvPr id. */
export function buildGroup(
  grp: GroupOptions,
  id: number,
  ctx: WriteContext,
): { xml: string; nextId: number } {
  const grpSpPrXml = groupShapePropertiesDesc.stringify(grp.grpSpPr, ctx) ?? "";
  let childId = id + 1;
  const children: string[] = [];
  for (const childShape of grp.shapes ?? []) {
    children.push(
      buildShapeContent(
        childShape,
        childId,
        `Shape ${childId}`,
        childShape.spPr,
        childShape.textBody,
        ctx,
        macroTextlinkAttrs(childShape),
      ),
    );
    childId++;
  }
  for (const childConn of grp.connectors ?? []) {
    children.push(
      buildConnectorContent(
        childConn,
        childId,
        `Connector ${childId}`,
        childConn.spPr,
        ctx,
        macroTextlinkAttrs(childConn),
        childConn,
      ),
    );
    childId++;
  }
  const xml =
    `<grpSp><nvGrpSpPr>${stringifyNonVisualDrawingProperties("cNvPr", id, grp, `Group ${id}`)}<cNvGrpSpPr/></nvGrpSpPr>` +
    `<grpSpPr>${grpSpPrXml}</grpSpPr>${children.join("")}</grpSp>`;
  return { xml, nextId: childId };
}

/** CT_Shape attribute string (macro/textlink) with leading space, or empty. */
function macroTextlinkAttrs(shape: { macro?: string; textlink?: string }): string {
  const a: string[] = [];
  if (shape.macro !== undefined) a.push(`macro="${escapeXml(shape.macro)}"`);
  if (shape.textlink !== undefined) a.push(`textlink="${escapeXml(shape.textlink)}"`);
  return a.length ? " " + a.join(" ") : "";
}
