import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { slideLayoutDesc } from "./slide-layout";
import type { SlideLayoutDescriptorOptions } from "./slide-layout";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: SlideLayoutDescriptorOptions) {
  const xml = slideLayoutDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return slideLayoutDesc.parse(el, readCtx);
}

describe("slideLayoutDesc round-trip", () => {
  it("round-trips layout with blank type from name", () => {
    const opts: SlideLayoutDescriptorOptions = {
      layout:
        '<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="Blank"/></p:sldLayout>',
    };
    const result = roundTrip(opts);
    expect(result.type).toBe("blank");
  });

  it("round-trips layout with title type from name", () => {
    const opts: SlideLayoutDescriptorOptions = {
      layout:
        '<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="Title Slide"/></p:sldLayout>',
    };
    const result = roundTrip(opts);
    expect(result.type).toBe("title");
  });

  it("round-trips layout with unknown name", () => {
    const opts: SlideLayoutDescriptorOptions = {
      layout:
        '<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="Custom Layout"/></p:sldLayout>',
    };
    const result = roundTrip(opts);
    expect(result.type).toBe("Custom Layout");
  });

  it("round-trips layout without cSld", () => {
    const opts: SlideLayoutDescriptorOptions = {
      layout: '<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>',
    };
    const result = roundTrip(opts);
    expect(result.type).toBeUndefined();
  });
});
