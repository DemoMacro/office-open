import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { tableStylesDesc } from "./table-styles";
import type { TableStylesDescriptorOptions } from "./table-styles";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: TableStylesDescriptorOptions) {
  const xml = tableStylesDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return tableStylesDesc.parse(el, readCtx);
}

describe("tableStylesDesc round-trip", () => {
  it("round-trips default (empty) table styles", () => {
    const opts: TableStylesDescriptorOptions = {};
    const result = roundTrip(opts);
    expect(result.opts).toBeDefined();
    expect(result.opts!.defaultStyleId).toBe("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}");
  });

  it("round-trips preserves def style id in XML", () => {
    const opts: TableStylesDescriptorOptions = {};
    const xml = tableStylesDesc.stringify(opts, writeCtx)!;
    expect(xml).toContain("def=");
    expect(xml).toContain("tblStyleLst");
  });
});
