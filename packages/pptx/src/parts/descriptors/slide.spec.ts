import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { slideDesc } from "./slide";
import type { SlideDescriptorOptions } from "./slide";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: SlideDescriptorOptions) {
  const xml = slideDesc.stringify(opts, writeCtx);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return slideDesc.parse(el, readCtx);
}

describe("slideDesc round-trip", () => {
  it("round-trips an empty slide", () => {
    const opts: SlideDescriptorOptions = {};
    const result = roundTrip(opts);

    expect(result).toBeDefined();
  });

  it("round-trips showMasterSp=false", () => {
    const opts: SlideDescriptorOptions = {
      showMasterSp: false,
    };
    const result = roundTrip(opts);

    expect(result.showMasterSp).toBe(false);
  });

  it("round-trips showMasterPhAnim=false", () => {
    const opts: SlideDescriptorOptions = {
      showMasterPhAnim: false,
    };
    const result = roundTrip(opts);

    expect(result.showMasterPhAnim).toBe(false);
  });

  it("round-trips background with color (no transparency)", () => {
    const opts: SlideDescriptorOptions = {
      background: { color: "FF5733" },
    };
    const result = roundTrip(opts);

    expect(result.background).toBeDefined();
    expect(result.background!.color).toBe("FF5733");
  });

  // Note: background with transparency is lossy — stringifyBackground outputs
  // a:srgbClr directly under p:bgPr when transparency is set, but readBackground
  // expects a:solidFill wrapper. This is a known round-trip inconsistency.

  it("round-trips transition fade", () => {
    const opts: SlideDescriptorOptions = {
      transition: { type: "fade", speed: "medium" },
    };
    const result = roundTrip(opts);

    expect(result.transition).toBeDefined();
    expect(result.transition!.type).toBe("fade");
    expect(result.transition!.speed).toBe("medium");
  });

  it("round-trips transition wipe with advance settings", () => {
    const opts: SlideDescriptorOptions = {
      transition: {
        type: "wipe",
        speed: "fast",
        advanceOnClick: false,
        advanceAfterMs: 5000,
      },
    };
    const result = roundTrip(opts);

    expect(result.transition).toBeDefined();
    expect(result.transition!.type).toBe("wipe");
    expect(result.transition!.speed).toBe("fast");
    expect(result.transition!.advanceOnClick).toBe(false);
    expect(result.transition!.advanceAfterMs).toBe(5000);
  });

  it("round-trips transition dissolve", () => {
    const opts: SlideDescriptorOptions = {
      transition: { type: "dissolve", speed: "slow" },
    };
    const result = roundTrip(opts);

    expect(result.transition).toBeDefined();
    expect(result.transition!.type).toBe("dissolve");
    expect(result.transition!.speed).toBe("slow");
  });

  it("round-trips transition push", () => {
    const opts: SlideDescriptorOptions = {
      transition: { type: "push", advanceOnClick: true },
    };
    const result = roundTrip(opts);

    expect(result.transition).toBeDefined();
    expect(result.transition!.type).toBe("push");
    expect(result.transition!.advanceOnClick).toBe(true);
  });

  // Note: headerFooter, customerData, controls are lossy in round-trip.
  // - headerFooter: stringify does not emit p:hf element
  // - customerData/controls: stringify places them inside p:cSld,
  //   but parse searches from p:sld root via findChild (direct children only)
});
