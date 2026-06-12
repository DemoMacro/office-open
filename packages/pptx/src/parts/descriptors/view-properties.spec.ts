import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { viewPropsDesc } from "./view-properties";
import type { ViewPropertiesDescriptorOptions } from "./view-properties";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: ViewPropertiesDescriptorOptions) {
  const xml = viewPropsDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return viewPropsDesc.parse(el, readCtx);
}

describe("viewPropsDesc round-trip", () => {
  it("round-trips lastView", () => {
    const opts: ViewPropertiesDescriptorOptions = {
      lastView: "slideView",
    };
    const result = roundTrip(opts);
    expect(result.lastView).toBe("slideView");
  });

  it("round-trips showComments false", () => {
    const opts: ViewPropertiesDescriptorOptions = {
      showComments: false,
    };
    const result = roundTrip(opts);
    expect(result.showComments).toBe(false);
  });

  it("round-trips gridSpacing", () => {
    const opts: ViewPropertiesDescriptorOptions = {
      gridSpacing: { cx: 100000, cy: 100000 },
    };
    const result = roundTrip(opts);
    expect(result.gridSpacing?.cx).toBe(100000);
    expect(result.gridSpacing?.cy).toBe(100000);
  });

  it("round-trips lastView slideMasterView", () => {
    const opts: ViewPropertiesDescriptorOptions = {
      lastView: "slideMasterView",
    };
    const result = roundTrip(opts);
    expect(result.lastView).toBe("slideMasterView");
  });

  it("round-trips lastView notesView", () => {
    const opts: ViewPropertiesDescriptorOptions = {
      lastView: "notesView",
    };
    const result = roundTrip(opts);
    expect(result.lastView).toBe("notesView");
  });
});
