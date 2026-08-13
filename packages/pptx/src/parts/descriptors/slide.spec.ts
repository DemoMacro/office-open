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
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return slideDesc.parse(el, readCtx);
}

describe("slideDesc round-trip", () => {
  it("round-trips an empty slide", () => {
    const opts: SlideDescriptorOptions = {};
    const result = roundTrip(opts);

    expect(result).toBeDefined();
  });

  it("round-trips showMasterShapes=false", () => {
    const opts: SlideDescriptorOptions = {
      showMasterShapes: false,
    };
    const result = roundTrip(opts);

    expect(result.showMasterShapes).toBe(false);
  });

  it("round-trips showMasterPlaceholderAnimations=false", () => {
    const opts: SlideDescriptorOptions = {
      showMasterPlaceholderAnimations: false,
    };
    const result = roundTrip(opts);

    expect(result.showMasterPlaceholderAnimations).toBe(false);
  });

  it("round-trips background with solid fill", () => {
    const opts: SlideDescriptorOptions = {
      background: { fill: { type: "solid", color: "FF5733" } },
    };
    const result = roundTrip(opts);

    expect(result.background).toBeDefined();
    const fill = result.background!.fill as { type: string; color: { value: string } };
    expect(fill.type).toBe("solid");
    expect(fill.color.value).toBe("FF5733");
  });

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
        advanceAfterTime: 5000,
      },
    };
    const result = roundTrip(opts);

    expect(result.transition).toBeDefined();
    expect(result.transition!.type).toBe("wipe");
    expect(result.transition!.speed).toBe("fast");
    expect(result.transition!.advanceOnClick).toBe(false);
    expect(result.transition!.advanceAfterTime).toBe(5000);
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

  it("round-trips headerFooter (p:hf)", () => {
    const opts: SlideDescriptorOptions = {
      headerFooter: { slideNumber: true, footer: true, header: true, dateTime: true },
    };
    const result = roundTrip(opts);

    expect(result.headerFooter).toBeDefined();
    expect(result.headerFooter!.slideNumber).toBe(true);
    expect(result.headerFooter!.footer).toBe(true);
    expect(result.headerFooter!.header).toBe(true);
    expect(result.headerFooter!.dateTime).toBe(true);
  });

  // Note: customerData/controls are lossy in round-trip.
  // - customerData/controls: stringify places them inside p:cSld,
  //   but parse searches from p:sld root via findChild (direct children only)
});
