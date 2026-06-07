/**
 * SpreadsheetDrawing descriptors for declarative XML read/write.
 *
 * @module
 */

import { escapeXml, findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type { CustomDescriptor } from "../../descriptor";
import type { XdrMarkerOptions, XdrTwoCellAnchorOptions } from "./spreadsheet-drawing";

// ── Helper: read a text-only child element's numeric content ──

function readNumChild(el: XmlElement, tag: string): number {
  const child = findChild(el, tag);
  if (!child?.elements?.length) return 0;
  const n = Number(child.elements[0]?.text ?? "");
  return Number.isNaN(n) ? 0 : n;
}

// ── xdr:from / xdr:to marker descriptor ──

function stringifyMarker(tag: string, opts: XdrMarkerOptions): string {
  return (
    `<${tag}>` +
    `<xdr:col>${opts.col}</xdr:col>` +
    `<xdr:colOff>${opts.colOff}</xdr:colOff>` +
    `<xdr:row>${opts.row}</xdr:row>` +
    `<xdr:rowOff>${opts.rowOff}</xdr:rowOff>` +
    `</${tag}>`
  );
}

function readMarker(el: XmlElement): Partial<XdrMarkerOptions> {
  return {
    col: readNumChild(el, "xdr:col"),
    colOff: readNumChild(el, "xdr:colOff"),
    row: readNumChild(el, "xdr:row"),
    rowOff: readNumChild(el, "xdr:rowOff"),
  };
}

export const xdrMarkerDesc: CustomDescriptor<XdrMarkerOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    return stringifyMarker("xdr:from", opts);
  },
  parse(el, _ctx) {
    return readMarker(el);
  },
};

// ── xdr:twoCellAnchor descriptor ──

export const xdrTwoCellAnchorDesc: CustomDescriptor<XdrTwoCellAnchorOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const attrStr = opts.editAs ? ` editAs="${escapeXml(opts.editAs)}"` : "";
    return `<xdr:twoCellAnchor${attrStr}/>`;
  },
  parse(el, _ctx) {
    const result: Partial<XdrTwoCellAnchorOptions> = {};
    if (el.attributes?.["editAs"] !== undefined) {
      result.editAs = String(el.attributes["editAs"]) as XdrTwoCellAnchorOptions["editAs"];
    }
    return result;
  },
};
