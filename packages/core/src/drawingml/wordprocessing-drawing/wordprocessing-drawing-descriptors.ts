/**
 * WordprocessingDrawing descriptors for DrawingML wp: namespace elements.
 *
 * @module
 */

import { escapeXml, findChild } from "@office-open/xml";

import type { CustomDescriptor } from "../../descriptor";
import type { WpNonVisualDrawingPropsOptions } from "./wordprocessing-drawing";

// ── wp:cNvPr descriptor ──

export const wpCNvPrDesc: CustomDescriptor<WpNonVisualDrawingPropsOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const attrParts: string[] = [];
    if (opts.id !== undefined) attrParts.push(`id="${opts.id}"`);
    if (opts.name) attrParts.push(`name="${escapeXml(opts.name)}"`);
    if (opts.descr) attrParts.push(`descr="${escapeXml(opts.descr)}"`);
    if (opts.hidden !== undefined) attrParts.push(`hidden="${opts.hidden ? 1 : 0}"`);
    if (attrParts.length === 0) return undefined;
    return `<wp:cNvPr ${attrParts.join(" ")}/>`;
  },
  parse(el, _ctx) {
    const result: Partial<WpNonVisualDrawingPropsOptions> = {};
    if (el.attributes) {
      if (el.attributes["id"] !== undefined) result.id = Number(el.attributes["id"]);
      if (el.attributes["name"] !== undefined) result.name = String(el.attributes["name"]);
      if (el.attributes["descr"] !== undefined) result.descr = String(el.attributes["descr"]);
      if (el.attributes["hidden"] !== undefined)
        result.hidden = el.attributes["hidden"] === 1 || el.attributes["hidden"] === "1";
    }
    return result;
  },
};

// ── wp:extent descriptor ──

export interface WpExtentOptions {
  cx?: number;
  cy?: number;
}

export const wpExtentDesc: CustomDescriptor<WpExtentOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    return `<wp:extent cx="${opts.cx}" cy="${opts.cy}"/>`;
  },
  parse(el, _ctx) {
    const result: Partial<WpExtentOptions> = {};
    if (el.attributes) {
      if (el.attributes["cx"] !== undefined) result.cx = Number(el.attributes["cx"]);
      if (el.attributes["cy"] !== undefined) result.cy = Number(el.attributes["cy"]);
    }
    return result;
  },
};

// ── wp:inline descriptor ──

export interface WpInlineOptions {
  extent?: WpExtentOptions;
  docPr?: WpNonVisualDrawingPropsOptions;
}

export const wpInlineDesc: CustomDescriptor<WpInlineOptions> = {
  kind: "custom",
  stringify(opts, ctx) {
    const parts: string[] = [];
    if (opts.extent) {
      const ext = wpExtentDesc.stringify(opts.extent, ctx);
      if (ext) parts.push(ext);
    }
    if (opts.docPr) {
      const pr = wpCNvPrDesc.stringify(opts.docPr, ctx);
      if (pr) parts.push(pr);
    }
    if (parts.length === 0) return "<wp:inline/>";
    return `<wp:inline>${parts.join("")}</wp:inline>`;
  },
  parse(el, _ctx) {
    const result: Partial<WpInlineOptions> = {};
    const ext = findChild(el, "wp:extent");
    if (ext) result.extent = wpExtentDesc.parse(ext, _ctx) as WpExtentOptions;
    const docPr = findChild(el, "wp:cNvPr");
    if (docPr) result.docPr = wpCNvPrDesc.parse(docPr, _ctx) as WpNonVisualDrawingPropsOptions;
    return result;
  },
};

// ── wp:anchor descriptor ──

export interface WpAnchorOptions {
  distT?: number;
  distB?: number;
  distL?: number;
  distR?: number;
  simplePos?: boolean;
  relativeHeight?: number;
  behindDoc?: boolean;
  locked?: boolean;
  layoutInCell?: boolean;
  allowOverlap?: boolean;
  extent?: WpExtentOptions;
  docPr?: WpNonVisualDrawingPropsOptions;
}

export const wpAnchorDesc: CustomDescriptor<WpAnchorOptions> = {
  kind: "custom",
  stringify(opts, ctx) {
    const attrParts: string[] = [];
    if (opts.distT !== undefined) attrParts.push(`distT="${opts.distT}"`);
    if (opts.distB !== undefined) attrParts.push(`distB="${opts.distB}"`);
    if (opts.distL !== undefined) attrParts.push(`distL="${opts.distL}"`);
    if (opts.distR !== undefined) attrParts.push(`distR="${opts.distR}"`);
    if (opts.simplePos !== undefined) attrParts.push(`simplePos="${opts.simplePos ? 1 : 0}"`);
    if (opts.relativeHeight !== undefined)
      attrParts.push(`relativeHeight="${opts.relativeHeight}"`);
    if (opts.behindDoc !== undefined) attrParts.push(`behindDoc="${opts.behindDoc ? 1 : 0}"`);
    if (opts.locked !== undefined) attrParts.push(`locked="${opts.locked ? 1 : 0}"`);
    if (opts.layoutInCell !== undefined)
      attrParts.push(`layoutInCell="${opts.layoutInCell ? 1 : 0}"`);
    if (opts.allowOverlap !== undefined)
      attrParts.push(`allowOverlap="${opts.allowOverlap ? 1 : 0}"`);
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";

    const parts: string[] = [];
    if (opts.extent) {
      const ext = wpExtentDesc.stringify(opts.extent, ctx);
      if (ext) parts.push(ext);
    }
    if (opts.docPr) {
      const pr = wpCNvPrDesc.stringify(opts.docPr, ctx);
      if (pr) parts.push(pr);
    }

    if (parts.length === 0 && !attrStr) return "<wp:anchor/>";
    if (parts.length === 0) return `<wp:anchor${attrStr}/>`;
    return `<wp:anchor${attrStr}>${parts.join("")}</wp:anchor>`;
  },
  parse(el, _ctx) {
    const result: Partial<WpAnchorOptions> = {};
    if (el.attributes) {
      if (el.attributes["distT"] !== undefined) result.distT = Number(el.attributes["distT"]);
      if (el.attributes["distB"] !== undefined) result.distB = Number(el.attributes["distB"]);
      if (el.attributes["distL"] !== undefined) result.distL = Number(el.attributes["distL"]);
      if (el.attributes["distR"] !== undefined) result.distR = Number(el.attributes["distR"]);
      if (el.attributes["simplePos"] !== undefined)
        result.simplePos = el.attributes["simplePos"] === 1 || el.attributes["simplePos"] === "1";
      if (el.attributes["relativeHeight"] !== undefined)
        result.relativeHeight = Number(el.attributes["relativeHeight"]);
      if (el.attributes["behindDoc"] !== undefined)
        result.behindDoc = el.attributes["behindDoc"] === 1 || el.attributes["behindDoc"] === "1";
      if (el.attributes["locked"] !== undefined)
        result.locked = el.attributes["locked"] === 1 || el.attributes["locked"] === "1";
      if (el.attributes["layoutInCell"] !== undefined)
        result.layoutInCell =
          el.attributes["layoutInCell"] === 1 || el.attributes["layoutInCell"] === "1";
      if (el.attributes["allowOverlap"] !== undefined)
        result.allowOverlap =
          el.attributes["allowOverlap"] === 1 || el.attributes["allowOverlap"] === "1";
    }
    const ext = findChild(el, "wp:extent");
    if (ext) result.extent = wpExtentDesc.parse(ext, _ctx) as WpExtentOptions;
    const docPr = findChild(el, "wp:cNvPr");
    if (docPr) result.docPr = wpCNvPrDesc.parse(docPr, _ctx) as WpNonVisualDrawingPropsOptions;
    return result;
  },
};
