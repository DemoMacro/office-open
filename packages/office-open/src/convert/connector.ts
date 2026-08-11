/**
 * Cross-format connector conversion.
 *
 * Connectors convert between pptx (p:cxnSp) and xlsx (xdr:cxnSp), which both
 * model a connector as a line geometry (spPr) plus optional endpoint glue
 * (startConnection/endConnection) and locks. The two endpoints (x1/y1/x2/y2 on
 * pptx) map to an xlsx bounding box (spPr.xfrm) with flip flags encoding the
 * direction the line was drawn — see ./position.
 *
 * docx has no standalone connector element (Word embeds connectors as wps
 * shapes flagged with a connector marker), so converting a connector to docx
 * is a no-op: the function warns and returns undefined so callers can skip it.
 *
 * @module
 */

import type { UniversalMeasure } from "@office-open/core";
import type { ShapePropertiesOptions } from "@office-open/core/drawingml";
import type { ConnectorShapeOptions as PptxConnectorOptions } from "@office-open/pptx";
import type { ConnectorOptions as XlsxConnectorOptions } from "@office-open/xlsx";

import { boxFromXlsxAnchor, boxToXlsx, toEmu } from "./position";
import type { AbsoluteBox } from "./position";

/** pptx two endpoints → absolute box; flip flags encode the draw direction. */
export function endpointsToBox(
  x1: number | UniversalMeasure | undefined,
  y1: number | UniversalMeasure | undefined,
  x2: number | UniversalMeasure | undefined,
  y2: number | UniversalMeasure | undefined,
): AbsoluteBox {
  const ax1 = toEmu(x1);
  const ay1 = toEmu(y1);
  const ax2 = toEmu(x2);
  const ay2 = toEmu(y2);
  return {
    x: Math.min(ax1, ax2),
    y: Math.min(ay1, ay2),
    width: Math.abs(ax2 - ax1),
    height: Math.abs(ay2 - ay1),
    ...(ax2 < ax1 ? { flipHorizontal: true } : {}),
    ...(ay2 < ay1 ? { flipVertical: true } : {}),
  };
}

/** absolute box → pptx two endpoints (restores direction from flip flags). */
export function boxToEndpoints(box: AbsoluteBox): {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
} {
  return {
    x1: box.flipHorizontal ? box.x + box.width : box.x,
    x2: box.flipHorizontal ? box.x : box.x + box.width,
    y1: box.flipVertical ? box.y + box.height : box.y,
    y2: box.flipVertical ? box.y : box.y + box.height,
  };
}

// ── → docx (no-op) ──

/**
 * docx has no standalone connector; warn and return undefined so callers skip
 * it. (Word embeds connectors as wps shapes; that path is not auto-derived
 * here.)
 */
export function toDocxConnector(_source: PptxConnectorOptions | XlsxConnectorOptions): undefined {
  console.warn("Connector conversion to docx is unsupported (docx has no standalone connector).");
  return undefined;
}

// ── → pptx ──

/** Convert an xlsx connector to a pptx connector. */
export function toPptxConnector(source: XlsxConnectorOptions): PptxConnectorOptions {
  const spPr = source.spPr;
  const box = boxFromXlsxAnchor(
    source,
    spPr.width,
    spPr.height,
    spPr.rotation,
    spPr.flipHorizontal,
    spPr.flipVertical,
  );
  const { x1, y1, x2, y2 } = boxToEndpoints(box);
  return {
    x1,
    y1,
    x2,
    y2,
    ...(spPr.outline !== undefined ? { outline: spPr.outline } : {}),
    ...(spPr.fill !== undefined ? { fill: spPr.fill } : {}),
    ...(source.locking ? { locking: source.locking } : {}),
    ...(source.startConnection ? { startConnection: source.startConnection } : {}),
    ...(source.endConnection ? { endConnection: source.endConnection } : {}),
    ...(source.name ? { name: source.name } : {}),
  };
}

// ── → xlsx ──

/** Convert a pptx connector to an xlsx connector. */
export function toXlsxConnector(source: PptxConnectorOptions): XlsxConnectorOptions {
  const box = endpointsToBox(source.x1, source.y1, source.x2, source.y2);
  const pos = boxToXlsx(box);
  const spPr: ShapePropertiesOptions = {
    x: pos.xfrmX,
    y: pos.xfrmY,
    width: box.width,
    height: box.height,
    // A connector renders as a line; carry the preset so xlsx emits prstGeom="line".
    geometry: "line",
    ...(source.outline !== undefined ? { outline: source.outline } : {}),
    ...(source.fill !== undefined ? { fill: source.fill } : {}),
    ...(box.flipHorizontal ? { flipHorizontal: true } : {}),
    ...(box.flipVertical ? { flipVertical: true } : {}),
  };
  return {
    ...pos.anchor,
    spPr,
    ...(source.locking ? { locking: source.locking } : {}),
    ...(source.startConnection ? { startConnection: source.startConnection } : {}),
    ...(source.endConnection ? { endConnection: source.endConnection } : {}),
    ...(source.name ? { name: source.name } : {}),
  };
}
