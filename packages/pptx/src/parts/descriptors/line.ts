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
  effectListDesc,
  fillDesc,
  findFillChild,
  outlineDesc,
  parseEndpointConnection,
  presetGeometryDesc,
  scene3DDesc,
  shape3DDesc,
  shapeLockingDesc,
  stringifyEndpointConnection,
  stringifyNonVisualDrawingProperties,
  textBodyDesc,
} from "@office-open/core/drawing";
import { attrBool, attrNum, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import type { ConnectorOptions, LineShapeOptions } from "@shared/shape/line-shape";
import { readShapeStyle } from "@shared/shape/shape";

import type { PptxWriteContext } from "../../context";
import { readCnvPr } from "./shape";
import { stringifyShapeStyle } from "./shape";
import { nextSlideDrawingId } from "./slide-drawing-ids";

// ── ID counters ──

let _nextLineId = 2;
let _nextConnectorId = 2;

// ── Shared endpoint-model spPr helpers (line + connector) ──

/** a:xfrm + a:prstGeom for an endpoint-model line (flip encodes direction). */
function stringifyLineXfrmGeometry(
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  geomXml = '<a:prstGeom prst="line"><a:avLst/></a:prstGeom>',
  rotation?: number,
): string {
  const attrs = [
    rotation !== undefined ? ` rot="${Math.round(rotation * 60000)}"` : "",
    x1 > x2 ? ' flipH="1"' : "",
    y1 > y2 ? ' flipV="1"' : "",
  ].join("");
  const offX = Math.min(x1, x2);
  const offY = Math.min(y1, y2);
  return (
    `<a:xfrm${attrs}><a:off x="${offX}" y="${offY}"/>` +
    `<a:ext cx="${Math.abs(x2 - x1)}" cy="${Math.abs(y2 - y1)}"/></a:xfrm>` +
    geomXml
  );
}

/** Parse endpoints (off/ext + flip) + paint from p:spPr of a line/connector. */
function parseLineSpPr(
  spPr: Element,
  ctx: ReadContext,
): Pick<LineShapeOptions, "x1" | "y1" | "x2" | "y2"> & {
  properties: NonNullable<LineShapeOptions["properties"]>;
  rotation?: number;
} {
  const result: ReturnType<typeof parseLineSpPr> = { properties: {} };

  const xfrm = findChild(spPr, "a:xfrm");
  if (xfrm) {
    const off = findChild(xfrm, "a:off");
    const ext = findChild(xfrm, "a:ext");
    const flipH = attrBool(xfrm, "flipH");
    const flipV = attrBool(xfrm, "flipV");
    const rotation = attrNum(xfrm, "rot");
    if (rotation !== undefined) result.rotation = rotation / 60000;

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

  const paint = result.properties;

  // Only parse fill when a fill child exists — fillDesc returns
  // { type: "none" } for an empty spPr, which would spuriously emit <a:noFill/>.
  const fillChild = findFillChild(spPr);
  if (fillChild) paint.fill = parse(fillDesc, fillChild, ctx);
  const ln = findChild(spPr, "a:ln");
  if (ln) paint.outline = parse(outlineDesc, ln, ctx);
  const effectLst = findChild(spPr, "a:effectLst");
  if (effectLst) paint.effects = parse(effectListDesc, effectLst, ctx);
  const scene3d = findChild(spPr, "a:scene3d");
  if (scene3d) paint.scene3d = scene3DDesc.parse(scene3d, ctx);
  const sp3d = findChild(spPr, "a:sp3d");
  if (sp3d) paint.shape3d = shape3DDesc.parse(sp3d, ctx);

  return result;
}

// ── LineShape (p:sp) descriptor ──

export const lineShapeDesc: CustomDescriptor<LineShapeOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const id = opts.id ?? nextSlideDrawingId(ctx) ?? _nextLineId++;
    const name = opts.name ?? `Line ${id}`;
    (ctx as PptxWriteContext).registerShapeId(name, id);

    const x1 = convertToEmu(opts.x1 ?? 0);
    const y1 = convertToEmu(opts.y1 ?? 0);
    const x2 = convertToEmu(opts.x2 ?? "100px");
    const y2 = convertToEmu(opts.y2 ?? "100px");

    const parts: string[] = [];

    // p:nvSpPr
    const spLocks = opts.locking ? (shapeLockingDesc.stringify(opts.locking, ctx) ?? "") : "";
    const txBoxAttr = opts.textBox ? ' txBox="1"' : "";
    const nvPr = opts.userDrawn ? '<p:nvPr userDrawn="1"/>' : "<p:nvPr/>";
    const cNvSpPr = spLocks
      ? `<p:cNvSpPr${txBoxAttr}>${spLocks}</p:cNvSpPr>`
      : `<p:cNvSpPr${txBoxAttr}/>`;
    parts.push(
      `<p:nvSpPr>${stringifyNonVisualDrawingProperties("p:cNvPr", id, opts, name)}${cNvSpPr}${nvPr}</p:nvSpPr>`,
    );

    // p:spPr
    const sp = opts.properties ?? {};
    const spPrParts: string[] = [];
    const spPrAttrs = opts.blackWhiteMode ? ` bwMode="${opts.blackWhiteMode}"` : "";
    spPrParts.push(stringifyLineXfrmGeometry(x1, y1, x2, y2, undefined, opts.rotation));

    // Fill
    if (sp.fill !== undefined) {
      const fillXml = stringify(fillDesc, sp.fill, ctx);
      if (fillXml) spPrParts.push(fillXml);
    }

    // Outline
    if (sp.outline) {
      const outlineXml = stringify(outlineDesc, sp.outline, ctx);
      if (outlineXml) spPrParts.push(outlineXml);
    }

    // Effects (a:effectLst after a:ln per CT_ShapeProperties)
    if (sp.effects) spPrParts.push(stringify(effectListDesc, sp.effects, ctx) ?? "");
    if (sp.scene3d) spPrParts.push(scene3DDesc.stringify(sp.scene3d, ctx) ?? "");
    if (sp.shape3d) spPrParts.push(shape3DDesc.stringify(sp.shape3d, ctx) ?? "");

    parts.push(`<p:spPr${spPrAttrs}>${spPrParts.join("")}</p:spPr>`);

    // p:style
    if (opts.style) {
      const styleXml = stringifyShapeStyle(opts.style, ctx);
      if (styleXml) parts.push(styleXml);
    }

    // p:txBody — source lines carry wrap/anchor hints; fresh ones keep the default.
    const txBodyContent = opts.textBody
      ? (textBodyDesc.stringify(opts.textBody, ctx) ?? "")
      : '<a:bodyPr wrap="square"/><a:lstStyle/><a:p/>';
    parts.push(`<p:txBody>${txBodyContent}</p:txBody>`);

    return `<p:sp>${parts.join("")}</p:sp>`;
  },

  parse(el, _ctx) {
    const result: Partial<LineShapeOptions> = {};

    // p:nvSpPr → id, name
    Object.assign(result, readCnvPr(el, "p:nvSpPr"));
    const nvSpPr = findChild(el, "p:nvSpPr");
    const cNvSpPr = nvSpPr ? findChild(nvSpPr, "p:cNvSpPr") : undefined;
    const spLocks = cNvSpPr ? findChild(cNvSpPr, "a:spLocks") : undefined;
    if (spLocks) {
      const locks = shapeLockingDesc.parse(spLocks, _ctx);
      if (locks && Object.keys(locks).length > 0) result.locking = locks;
    }
    if (attrBool(cNvSpPr, "txBox")) result.textBox = true;
    if (attrBool(nvSpPr ? findChild(nvSpPr, "p:nvPr") : undefined, "userDrawn"))
      result.userDrawn = true;

    // p:spPr → endpoints (off/ext + flip) + fill/outline/effects
    const spPr = findChild(el, "p:spPr");
    if (spPr) {
      Object.assign(result, parseLineSpPr(spPr, _ctx));
      const bwMode = spPr.attributes?.["bwMode"];
      if (bwMode !== undefined)
        result.blackWhiteMode = bwMode as LineShapeOptions["blackWhiteMode"];
    }

    // p:style
    const lineStyle = findChild(el, "p:style");
    if (lineStyle) result.style = readShapeStyle(lineStyle, _ctx);

    // p:txBody — keep the source's wrap/anchor hints and empty-paragraph rPr.
    const txBody = findChild(el, "p:txBody");
    if (txBody) result.textBody = textBodyDesc.parse(txBody, _ctx);

    return result as LineShapeOptions;
  },
};

// ── ConnectorShape (p:cxnSp) descriptor ──

export const connectorShapeDesc: CustomDescriptor<ConnectorOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const id = opts.id ?? nextSlideDrawingId(ctx) ?? _nextConnectorId++;
    const name = opts.name ?? `Connector ${id}`;
    (ctx as PptxWriteContext).registerShapeId(name, id);

    const x1 = convertToEmu(opts.x1 ?? 0);
    const y1 = convertToEmu(opts.y1 ?? 0);
    const x2 = convertToEmu(opts.x2 ?? "100px");
    const y2 = convertToEmu(opts.y2 ?? "100px");

    // Preset geometry: bent/elbow connectors carry their own form + guides.
    const geomXml =
      opts.properties?.geometry !== undefined
        ? (presetGeometryDesc.stringify(
            typeof opts.properties.geometry === "string"
              ? { preset: opts.properties.geometry }
              : opts.properties.geometry,
            ctx,
          ) ?? "")
        : undefined;

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
      `<p:nvCxnSpPr>${stringifyNonVisualDrawingProperties("p:cNvPr", id, opts, name)}${cNvCxnSpPr}${opts.userDrawn ? '<p:nvPr userDrawn="1"/>' : "<p:nvPr/>"}</p:nvCxnSpPr>`,
    );

    // p:spPr
    const sp = opts.properties ?? {};
    const spPrParts: string[] = [];
    const spPrAttrs = opts.blackWhiteMode ? ` bwMode="${opts.blackWhiteMode}"` : "";
    spPrParts.push(stringifyLineXfrmGeometry(x1, y1, x2, y2, geomXml, opts.rotation));

    // Fill
    if (sp.fill !== undefined) {
      const fillXml = stringify(fillDesc, sp.fill, ctx);
      if (fillXml) spPrParts.push(fillXml);
    }

    // Outline (arrowheads live inside the outline as headEnd/tailEnd)
    if (sp.outline) {
      const outlineXml = stringify(outlineDesc, sp.outline, ctx);
      if (outlineXml) spPrParts.push(outlineXml);
    }

    // Effects (a:effectLst after a:ln per CT_ShapeProperties)
    if (sp.effects) spPrParts.push(stringify(effectListDesc, sp.effects, ctx) ?? "");
    if (sp.scene3d) spPrParts.push(scene3DDesc.stringify(sp.scene3d, ctx) ?? "");
    if (sp.shape3d) spPrParts.push(shape3DDesc.stringify(sp.shape3d, ctx) ?? "");

    parts.push(`<p:spPr${spPrAttrs}>${spPrParts.join("")}</p:spPr>`);

    // p:style
    if (opts.style) {
      const styleXml = stringifyShapeStyle(opts.style, ctx);
      if (styleXml) parts.push(styleXml);
    }

    return `<p:cxnSp>${parts.join("")}</p:cxnSp>`;
  },

  parse(el, _ctx) {
    const result: Partial<ConnectorOptions> = {};

    // p:nvCxnSpPr → id, name, and optional cNvCxnSpPr (locks + connections)
    const nvCxnSpPr = findChild(el, "p:nvCxnSpPr");
    if (nvCxnSpPr) {
      Object.assign(result, readCnvPr(nvCxnSpPr));
      if (attrBool(findChild(nvCxnSpPr, "p:nvPr"), "userDrawn")) result.userDrawn = true;
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

    // p:spPr → endpoints (off/ext + flip) + geometry/fill/outline/effects
    const spPr = findChild(el, "p:spPr");
    if (spPr) {
      const { properties, ...endpoints } = parseLineSpPr(spPr, _ctx);
      Object.assign(result, endpoints);
      result.properties = properties;
      const bwMode = spPr.attributes?.["bwMode"];
      if (bwMode !== undefined)
        result.blackWhiteMode = bwMode as ConnectorOptions["blackWhiteMode"];
      // Non-line presets or adjusted guides must survive the round-trip.
      const prstGeom = findChild(spPr, "a:prstGeom");
      if (prstGeom) {
        const geom = presetGeometryDesc.parse(prstGeom, _ctx);
        if (geom.preset !== undefined && (geom.preset !== "line" || geom.adjustmentValues)) {
          result.properties.geometry = geom;
        }
      }
    }

    // p:style
    const cxnStyle = findChild(el, "p:style");
    if (cxnStyle) result.style = readShapeStyle(cxnStyle, _ctx);

    return result as ConnectorOptions;
  },
};
