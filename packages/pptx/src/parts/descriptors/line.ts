/**
 * Line and Connector descriptors for PPTX.
 *
 * @module
 */

import { convertEmuToPixels, convertPixelsToEmu, xsdLineEndSize } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { parse, stringify } from "@office-open/core/descriptor";
import { fillDesc, outlineDesc } from "@office-open/core/drawingml";
import { attr, attrBool, attrNum, findChild } from "@office-open/xml";
import { escapeXml } from "@office-open/xml";
import type { FillOptions } from "@shared/drawingml/fill";
import type { OutlineOptions } from "@shared/drawingml/outline";

import { toCoreOutlineOptions, readOutlineCompat } from "./shape";

// ── Types ──

export interface LineShapeDescriptorOptions {
  id?: number;
  name?: string;
  x1?: number;
  y1?: number;
  x2?: number;
  y2?: number;
  fill?: FillOptions;
  outline?: OutlineOptions;
}

export type ArrowheadType = "triangle" | "stealth" | "diamond" | "oval" | "open" | "none";

export interface ConnectorShapeDescriptorOptions {
  id?: number;
  name?: string;
  x1?: number;
  y1?: number;
  x2?: number;
  y2?: number;
  fill?: FillOptions;
  outline?: OutlineOptions;
  beginArrowhead?: ArrowheadType;
  endArrowhead?: ArrowheadType;
  arrowheadWidth?: "small" | "medium" | "large";
  arrowheadLength?: "small" | "medium" | "large";
}

// ── ID counters ──

let _nextLineId = 2;
let _nextConnectorId = 2;

// ── LineShape (p:sp) descriptor ──

export const lineShapeDesc: CustomDescriptor<LineShapeDescriptorOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const id = opts.id ?? _nextLineId++;
    const name = opts.name ?? `Line ${id}`;

    const x1 = convertPixelsToEmu(opts.x1 ?? 0);
    const y1 = convertPixelsToEmu(opts.y1 ?? 0);
    const x2 = convertPixelsToEmu(opts.x2 ?? 100);
    const y2 = convertPixelsToEmu(opts.y2 ?? 100);

    const offX = Math.min(x1, x2);
    const offY = Math.min(y1, y2);

    const parts: string[] = [];

    // p:nvSpPr
    parts.push(
      `<p:nvSpPr><p:cNvPr id="${id}" name="${escapeXml(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>`,
    );

    // p:spPr
    const spPrParts: string[] = [];

    // a:xfrm (with optional flip)
    const xfrmAttrs: string[] = [];
    if (x1 > x2) xfrmAttrs.push(' flipH="1"');
    if (y1 > y2) xfrmAttrs.push(' flipV="1"');

    spPrParts.push(
      `<a:xfrm${xfrmAttrs.join("")}><a:off x="${offX}" y="${offY}"/><a:ext cx="${Math.abs(x2 - x1)}" cy="${Math.abs(y2 - y1)}"/></a:xfrm>`,
    );

    // PresetGeometry
    spPrParts.push('<a:prstGeom prst="line"><a:avLst/></a:prstGeom>');

    // Fill
    if (opts.fill !== undefined) {
      const fillXml = stringify(fillDesc, opts.fill, ctx);
      if (fillXml) spPrParts.push(fillXml);
    }

    // Outline
    if (opts.outline) {
      const coreOutline = toCoreOutlineOptions(opts.outline);
      const outlineXml = stringify(outlineDesc, coreOutline, ctx);
      if (outlineXml) spPrParts.push(outlineXml);
    }

    parts.push(`<p:spPr>${spPrParts.join("")}</p:spPr>`);

    // p:txBody
    parts.push('<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p/></p:txBody>');

    return `<p:sp>${parts.join("")}</p:sp>`;
  },

  parse(el, _ctx) {
    const result: Record<string, unknown> = {};

    // p:nvSpPr → id, name
    const nvSpPr = findChild(el, "p:nvSpPr");
    if (nvSpPr) {
      const cNvPr = findChild(nvSpPr, "p:cNvPr");
      if (cNvPr) {
        const id = attrNum(cNvPr, "id");
        if (id !== undefined) result.id = id;
        const name = attr(cNvPr, "name");
        if (name) result.name = name;
      }
    }

    // p:spPr → endpoints (off/ext + flip)
    const spPr = findChild(el, "p:spPr");
    if (spPr) {
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

          const x1 = flipH ? offX + cx : offX;
          const y1 = flipV ? offY + cy : offY;
          const x2 = flipH ? offX : offX + cx;
          const y2 = flipV ? offY : offY + cy;

          result.x1 = convertEmuToPixels(x1);
          result.y1 = convertEmuToPixels(y1);
          result.x2 = convertEmuToPixels(x2);
          result.y2 = convertEmuToPixels(y2);
        }
      }

      const fillResult = parse(fillDesc, spPr, _ctx);
      if (fillResult && Object.keys(fillResult).length > 0) result.fill = fillResult;
      const ln = findChild(spPr, "a:ln");
      if (ln) result.outline = readOutlineCompat(ln);
    }

    return result as unknown as LineShapeDescriptorOptions;
  },
};

// ── Arrowhead mapping ──

const ARROWHEAD_MAP: Record<string, string> = {
  triangle: "triangle",
  stealth: "stealth",
  diamond: "diamond",
  oval: "oval",
  open: "arrow",
  none: "none",
};

// ── ConnectorShape (p:cxnSp) descriptor ──

export const connectorShapeDesc: CustomDescriptor<ConnectorShapeDescriptorOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const id = opts.id ?? _nextConnectorId++;
    const name = opts.name ?? `Connector ${id}`;

    const x1 = convertPixelsToEmu(opts.x1 ?? 0);
    const y1 = convertPixelsToEmu(opts.y1 ?? 0);
    const x2 = convertPixelsToEmu(opts.x2 ?? 100);
    const y2 = convertPixelsToEmu(opts.y2 ?? 100);

    const offX = Math.min(x1, x2);
    const offY = Math.min(y1, y2);

    const parts: string[] = [];

    // p:nvCxnSpPr
    parts.push(
      `<p:nvCxnSpPr><p:cNvPr id="${id}" name="${escapeXml(name)}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>`,
    );

    // p:spPr
    const spPrParts: string[] = [];

    // a:xfrm
    const xfrmAttrs: string[] = [];
    if (x1 > x2) xfrmAttrs.push(' flipH="1"');
    if (y1 > y2) xfrmAttrs.push(' flipV="1"');

    spPrParts.push(
      `<a:xfrm${xfrmAttrs.join("")}><a:off x="${offX}" y="${offY}"/><a:ext cx="${Math.abs(x2 - x1)}" cy="${Math.abs(y2 - y1)}"/></a:xfrm>`,
    );

    spPrParts.push('<a:prstGeom prst="line"><a:avLst/></a:prstGeom>');

    // Fill
    if (opts.fill !== undefined) {
      const fillXml = stringify(fillDesc, opts.fill, ctx);
      if (fillXml) spPrParts.push(fillXml);
    }

    // Outline with arrowheads
    const coreOutline = opts.outline ? toCoreOutlineOptions(opts.outline) : {};
    const hasArrowheads = opts.beginArrowhead || opts.endArrowhead;

    if (opts.outline || hasArrowheads) {
      if (hasArrowheads) {
        if (opts.beginArrowhead) {
          (coreOutline as Record<string, unknown>).tailEnd = {
            type: ARROWHEAD_MAP[opts.beginArrowhead] ?? "none",
            ...(opts.arrowheadWidth ? { w: opts.arrowheadWidth } : {}),
            ...(opts.arrowheadLength ? { len: opts.arrowheadLength } : {}),
          };
        }
        if (opts.endArrowhead) {
          (coreOutline as Record<string, unknown>).headEnd = {
            type: ARROWHEAD_MAP[opts.endArrowhead] ?? "none",
            ...(opts.arrowheadWidth ? { w: opts.arrowheadWidth } : {}),
            ...(opts.arrowheadLength ? { len: opts.arrowheadLength } : {}),
          };
        }
      }
      const outlineXml = stringify(outlineDesc, coreOutline, ctx);
      if (outlineXml) spPrParts.push(outlineXml);
    }

    parts.push(`<p:spPr>${spPrParts.join("")}</p:spPr>`);

    return `<p:cxnSp>${parts.join("")}</p:cxnSp>`;
  },

  parse(el, _ctx) {
    const result: Record<string, unknown> = {};

    // p:nvCxnSpPr → id, name
    const nvCxnSpPr = findChild(el, "p:nvCxnSpPr");
    if (nvCxnSpPr) {
      const cNvPr = findChild(nvCxnSpPr, "p:cNvPr");
      if (cNvPr) {
        const id = attrNum(cNvPr, "id");
        if (id !== undefined) result.id = id;
        const name = attr(cNvPr, "name");
        if (name) result.name = name;
      }
    }

    // p:spPr → endpoints (off/ext + flip)
    const spPr = findChild(el, "p:spPr");
    if (spPr) {
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

          const x1 = flipH ? offX + cx : offX;
          const y1 = flipV ? offY + cy : offY;
          const x2 = flipH ? offX : offX + cx;
          const y2 = flipV ? offY : offY + cy;

          result.x1 = convertEmuToPixels(x1);
          result.y1 = convertEmuToPixels(y1);
          result.x2 = convertEmuToPixels(x2);
          result.y2 = convertEmuToPixels(y2);
        }
      }

      const fillResult = parse(fillDesc, spPr, _ctx);
      if (fillResult && Object.keys(fillResult).length > 0) result.fill = fillResult;
      const ln = findChild(spPr, "a:ln");
      if (ln) result.outline = readOutlineCompat(ln);

      // Arrowheads from outline
      if (ln) {
        const headEnd = findChild(ln, "a:headEnd");
        if (headEnd) {
          const type = attr(headEnd, "type");
          if (type) result.endArrowhead = type as ArrowheadType;
          const w = attr(headEnd, "w");
          if (w) result.arrowheadWidth = xsdLineEndSize.from(w) as "small" | "medium" | "large";
          const len = attr(headEnd, "len");
          if (len)
            result.arrowheadLength = xsdLineEndSize.from(len) as "small" | "medium" | "large";
        }
        const tailEnd = findChild(ln, "a:tailEnd");
        if (tailEnd) {
          const type = attr(tailEnd, "type");
          if (type) result.beginArrowhead = type as ArrowheadType;
        }
      }
    }

    return result as unknown as ConnectorShapeDescriptorOptions;
  },
};
