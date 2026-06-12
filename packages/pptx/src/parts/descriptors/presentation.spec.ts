import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { presentationDesc } from "./presentation";
import type { PresentationDescriptorOptions } from "./presentation";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: PresentationDescriptorOptions) {
  const xml = presentationDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return presentationDesc.parse(el, readCtx);
}

describe("presentationDesc round-trip", () => {
  it("round-trips slide size", () => {
    const opts: PresentationDescriptorOptions = {
      slideWidth: 12192000,
      slideHeight: 6858000,
      slideIds: [256, 257],
      masterCount: 1,
    };
    const result = roundTrip(opts);
    expect(result.slideWidth).toBe(12192000);
    expect(result.slideHeight).toBe(6858000);
  });

  it("round-trips slideIds and masterCount", () => {
    const opts: PresentationDescriptorOptions = {
      slideIds: [256, 257, 258],
      masterCount: 2,
    };
    const result = roundTrip(opts);
    expect(result.slideIds).toEqual([256, 257, 258]);
    expect(result.masterCount).toBe(2);
  });

  it("round-trips rtl", () => {
    const opts: PresentationDescriptorOptions = {
      slideIds: [256],
      masterCount: 1,
      rtl: true,
    };
    const result = roundTrip(opts);
    expect(result.rtl).toBe(true);
  });

  it("round-trips autoCompressPictures false", () => {
    const opts: PresentationDescriptorOptions = {
      slideIds: [256],
      masterCount: 1,
      autoCompressPictures: false,
    };
    const result = roundTrip(opts);
    expect(result.autoCompressPictures).toBe(false);
  });

  it("round-trips firstSlideNum", () => {
    const opts: PresentationDescriptorOptions = {
      slideIds: [256],
      masterCount: 1,
      firstSlideNum: 5,
    };
    const result = roundTrip(opts);
    expect(result.firstSlideNum).toBe(5);
  });
});
