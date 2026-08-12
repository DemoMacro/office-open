import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import type { PresentationPropertiesOptions } from "@parts/presentation-properties";
import { describe, expect, it } from "vite-plus/test";

import { presentationPropertiesDesc } from "./presentation-properties";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: PresentationPropertiesOptions) {
  const xml = presentationPropertiesDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return presentationPropertiesDesc.parse(el, readCtx);
}

describe("presentationPropertiesDesc round-trip", () => {
  it("round-trips show with loop", () => {
    const opts: PresentationPropertiesOptions = {
      show: { loop: true },
    };
    const result = roundTrip(opts);
    expect(result.show?.loop).toBe(true);
  });

  it("round-trips show with type kiosk", () => {
    const opts: PresentationPropertiesOptions = {
      show: { type: "kiosk" },
    };
    const result = roundTrip(opts);
    expect(result.show?.type).toBe("kiosk");
  });

  it("round-trips show with useTimings", () => {
    const opts: PresentationPropertiesOptions = {
      show: { useTimings: true },
    };
    const result = roundTrip(opts);
    expect(result.show?.useTimings).toBe(true);
  });

  it("round-trips show with showNarration false", () => {
    const opts: PresentationPropertiesOptions = {
      show: { showNarration: false },
    };
    const result = roundTrip(opts);
    expect(result.show?.showNarration).toBe(false);
  });

  it("round-trips empty options", () => {
    const xml = presentationPropertiesDesc.stringify(
      {} as PresentationPropertiesOptions,
      writeCtx,
    )!;
    const doc = parseXml(xml);
    const el = doc.elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const result = presentationPropertiesDesc.parse(el, readCtx);
    expect(result).toBeDefined();
  });
});
