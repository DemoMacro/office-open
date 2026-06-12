import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { presPropsDesc } from "./presentation-properties";
import type { PresPropsDescriptorOptions } from "./presentation-properties";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: PresPropsDescriptorOptions) {
  const xml = presPropsDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return presPropsDesc.parse(el, readCtx);
}

describe("presPropsDesc round-trip", () => {
  it("round-trips show with loop", () => {
    const opts: PresPropsDescriptorOptions = {
      show: { loop: true },
    };
    const result = roundTrip(opts);
    expect(result.show?.loop).toBe(true);
  });

  it("round-trips show with type kiosk", () => {
    const opts: PresPropsDescriptorOptions = {
      show: { type: "kiosk" },
    };
    const result = roundTrip(opts);
    expect(result.show?.type).toBe("kiosk");
  });

  it("round-trips show with useTimings", () => {
    const opts: PresPropsDescriptorOptions = {
      show: { useTimings: true },
    };
    const result = roundTrip(opts);
    expect(result.show?.useTimings).toBe(true);
  });

  it("round-trips show with showNarration false", () => {
    const opts: PresPropsDescriptorOptions = {
      show: { showNarration: false },
    };
    const result = roundTrip(opts);
    expect(result.show?.showNarration).toBe(false);
  });

  it("round-trips empty options", () => {
    const xml = presPropsDesc.stringify({} as PresPropsDescriptorOptions, writeCtx)!;
    const doc = parseXml(xml);
    const el = doc.elements![0];
    const result = presPropsDesc.parse(el, readCtx);
    expect(result).toBeDefined();
  });
});
