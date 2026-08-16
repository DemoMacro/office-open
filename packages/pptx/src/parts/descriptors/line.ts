/**
 * Line and Connector descriptors for PPTX.
 *
 * @module
 */

import { convertToEmu } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { parse, stringify } from "@office-open/core/descriptor";
import type { ReadContext } from "@office-open/core/descriptor";
import {
  connectorLockingDesc,
  fillDesc,
  findFillChild,
  outlineDesc,
  parseEndpointConnection,
  stringifyEndpointConnection,
  stringifyNonVisualDrawingProperties,
  parseNonVisualDrawingProperties,
} from "@office-open/core/drawing";
import { attrBool, attrNum, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import type { ConnectorOptions, LineShapeOptions } from "@shared/shape/line-shape";

// ── ID counters ──

let _nextLineId = 2;
let _nextConnectorId = 2;

// ── Shared endpoint-model spPr helpers (line + connector) ──

/** a:xfrm + a:prstGeom for an endpoint-model line (flip encodes direction). */
function stringifyLineXfrmGeometry(x1: number, y1: number, x2: number, y2: number): string {
  const attrs = [x1 > x2 ? ' flipH="1"' : "", y1 > y2 ? ' flipV="1"' : ""].join("");
  const offX = Math.min(x1, x2);
  const offY = Math.min(y1, y2);
  return (
    `<a:xfrm${attrs}><a:off x="${offX}" y="${offY}"/>` +
    `<a:ext cx="${Math.abs(x2 - x1)}" cy="${Math.abs(y2 - y1)}"/></a:xfrm>` +
    `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>`
  );
}

/** Parse endpoints + fill/outline from p:spPr of a line/connector. */
function parseLineSpPr(
  spPr: Element,
  ctx: ReadContext,
): Pick<LineShapeOptions, "x1" | "y1" | "x2" | "y2" | "fill" | "outline"> {
  const result: ReturnType<typeof parseLineSpPr> = {};

  const xfrm = findChild(spPr, "a:xfrm");
  if (xfrm) {
    const off = findChild(xfrm, "a:off");
    const ext = findChild(xfrm, "a:ext");
    const flipH = attrBool(xfrm, "flipH");
    const flipV = attrBool(xfrm, "flipV");

    if (off && ext) {
      const offX = attrNum(off, "x") ?? 0;
      const offY = attrNum(off, "y") ?? 0;
      const cx = attrNum(ext, "cx") ?? 0;
      const cy = attrNum(ext, "cy") ?? 0;

      result.x1 = flipH ? offX + cx : offX;
      result.y1 = flipV ? offY + cy : offY;
      result.x2 = flipH ? offX : offX + cx;
      result.y2 = flipV ? offY : offY + cy;
    }
  }

  // Only parse fill when a fill child exists — fillDesc returns
  // { type: "none" } for an empty spPr, which would spuriously emit <a:noFill/>.
  const fillChild = findFillChild(spPr);
  if (fillChild) result.fill = parse(fillDesc, fillChild, ctx);
  const ln = findChild(spPr, "a:ln");
  if (ln) result.outline = parse(outlineDesc, ln, ctx);

  return result;
}

// ── LineShape (p:sp) descriptor ──

export const lineShapeDesc: CustomDescriptor<LineShapeOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const id = opts.id ?? _nextLineId++;
    const name = opts.name ?? `Line ${id}`;

    const x1 = convertToEmu(opts.x1 ?? 0);
    const y1 = convertToEmu(opts.y1 ?? 0);
    const x2 = convertToEmu(opts.x2 ?? "100px");
    const y2 = convertToEmu(opts.y2 ?? "100px");

    const parts: string[] = [];

    // p:nvSpPr
    parts.push(
      `<p:nvSpPr>${stringifyNonVisualDrawingProperties("p:cNvPr", id, opts, name)}<p:cNvSpPr/><p:nvPr/></p:nvSpPr>`,
    );

    // p:spPr
    const spPrParts: string[] = [];
    spPrParts.push(stringifyLineXfrmGeometry(x1, y1, x2, y2));

    // Fill
    if (opts.fill !== undefined) {
      const fillXml = stringify(fillDesc, opts.fill, ctx);
      if (fillXml) spPrParts.push(fillXml);
    }

    // Outline
    if (opts.outline) {
      const outlineXml = stringify(outlineDesc, opts.outline, ctx);
      if (outlineXml) spPrParts.push(outlineXml);
    }

    parts.push(`<p:spPr>${spPrParts.join("")}</p:spPr>`);

    // p:txBody
    parts.push('<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p/></p:txBody>');

    return `<p:sp>${parts.join("")}</p:sp>`;
  },

  parse(el, _ctx) {
    const result: Partial<LineShapeOptions> = {};

    // p:nvSpPr → id, name
    const nvSpPr = findChild(el, "p:nvSpPr");
    if (nvSpPr) {
      const cNvPr = findChild(nvSpPr, "p:cNvPr");
      if (cNvPr) {
        Object.assign(result, parseNonVisualDrawingProperties(cNvPr));
        const id = attrNum(cNvPr, "id");
        if (id !== undefined) result.id = id;
      }
    }

    // p:spPr → endpoints (off/ext + flip) + fill/outline
    const spPr = findChild(el, "p:spPr");
    if (spPr) Object.assign(result, parseLineSpPr(spPr, _ctx));

    return result as LineShapeOptions;
  },
};

// ── ConnectorShape (p:cxnSp) descriptor ──

export const connectorShapeDesc: CustomDescriptor<ConnectorOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const id = opts.id ?? _nextConnectorId++;
    const name = opts.name ?? `Connector ${id}`;

    const x1 = convertToEmu(opts.x1 ?? 0);
    const y1 = convertToEmu(opts.y1 ?? 0);
    const x2 = convertToEmu(opts.x2 ?? "100px");
    const y2 = convertToEmu(opts.y2 ?? "100px");

    const parts: string[] = [];

    // p:nvCxnSpPr — cNvCxnSpPr holds optional cxnSpLocks/stCxn/endCxn.
    const cNvCxnSpPrInner: string[] = [];
    if (opts.locking) {
      const locks = stringify(connectorLockingDesc, opts.locking, ctx);
      if (locks) cNvCxnSpPrInner.push(locks);
    }
    if (opts.startConnection) {
      cNvCxnSpPrInner.push(stringifyEndpointConnection("stCxn", opts.startConnection));
    }
    if (opts.endConnection) {
      cNvCxnSpPrInner.push(stringifyEndpointConnection("endCxn", opts.endConnection));
    }
    const cNvCxnSpPr = cNvCxnSpPrInner.length
      ? `<p:cNvCxnSpPr>${cNvCxnSpPrInner.join("")}</p:cNvCxnSpPr>`
      : "<p:cNvCxnSpPr/>";
    parts.push(
      `<p:nvCxnSpPr>${stringifyNonVisualDrawingProperties("p:cNvPr", id, opts, name)}${cNvCxnSpPr}<p:nvPr/></p:nvCxnSpPr>`,
    );

    // p:spPr
    const spPrParts: string[] = [];
    spPrParts.push(stringifyLineXfrmGeometry(x1, y1, x2, y2));

    // Fill
    if (opts.fill !== undefined) {
      const fillXml = stringify(fillDesc, opts.fill, ctx);
      if (fillXml) spPrParts.push(fillXml);
    }

    // Outline (arrowheads live inside the outline as headEnd/tailEnd)
    if (opts.outline) {
      const outlineXml = stringify(outlineDesc, opts.outline, ctx);
      if (outlineXml) spPrParts.push(outlineXml);
    }

    parts.push(`<p:spPr>${spPrParts.join("")}</p:spPr>`);

    return `<p:cxnSp>${parts.join("")}</p:cxnSp>`;
  },

  parse(el, _ctx) {
    const result: Partial<ConnectorOptions> = {};

    // p:nvCxnSpPr → id, name, and optional cNvCxnSpPr (locks + connections)
    const nvCxnSpPr = findChild(el, "p:nvCxnSpPr");
    if (nvCxnSpPr) {
      const cNvPr = findChild(nvCxnSpPr, "p:cNvPr");
      if (cNvPr) {
        Object.assign(result, parseNonVisualDrawingProperties(cNvPr));
        const id = attrNum(cNvPr, "id");
        if (id !== undefined) result.id = id;
      }
      const cNvCxnSpPr = findChild(nvCxnSpPr, "p:cNvCxnSpPr");
      if (cNvCxnSpPr) {
        const cxnSpLocks = findChild(cNvCxnSpPr, "a:cxnSpLocks");
        if (cxnSpLocks) {
          const locks = parse(connectorLockingDesc, cxnSpLocks, _ctx);
          if (locks && Object.keys(locks).length > 0) result.locking = locks;
        }
        const stCxn = findChild(cNvCxnSpPr, "a:stCxn");
        if (stCxn) {
          const conn = parseEndpointConnection(stCxn);
          if (conn) result.startConnection = conn;
        }
        const endCxn = findChild(cNvCxnSpPr, "a:endCxn");
        if (endCxn) {
          const conn = parseEndpointConnection(endCxn);
          if (conn) result.endConnection = conn;
        }
      }
    }

    // p:spPr → endpoints (off/ext + flip) + fill/outline
    const spPr = findChild(el, "p:spPr");
    if (spPr) Object.assign(result, parseLineSpPr(spPr, _ctx));

    return result as ConnectorOptions;
  },
};
