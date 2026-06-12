import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { backgroundDesc } from "./background";
import type { BackgroundDescriptorOptions } from "./background";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: BackgroundDescriptorOptions) {
  const xml = backgroundDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return backgroundDesc.parse(el, readCtx);
}

describe("backgroundDesc round-trip", () => {
  it("round-trips solidFill color", () => {
    const opts: BackgroundDescriptorOptions = {
      fill: { type: "solid", color: "FF0000" },
    };
    const result = roundTrip(opts);
    const fill = result.fill! as { type: string; color: { value: string } };

    expect(fill.type).toBe("solid");
    expect(fill.color.value).toBe("FF0000");
  });

  it("round-trips noFill", () => {
    const opts: BackgroundDescriptorOptions = {
      fill: { type: "none" },
    };
    const result = roundTrip(opts);
    const fill = result.fill! as { type: string };

    expect(fill.type).toBe("none");
  });

  it("round-trips shadeToTitle", () => {
    const opts: BackgroundDescriptorOptions = {
      fill: { type: "solid", color: "4472C4" },
      shadeToTitle: true,
    };
    const result = roundTrip(opts);

    expect(result.shadeToTitle).toBe(true);
  });

  it("round-trips blackWhiteMode", () => {
    const opts: BackgroundDescriptorOptions = {
      fill: { type: "solid", color: "000000" },
      blackWhiteMode: "gray",
    };
    const result = roundTrip(opts);

    expect(result.blackWhiteMode).toBe("gray");
  });

  it("round-trips gradientFill", () => {
    const opts: BackgroundDescriptorOptions = {
      fill: {
        type: "gradient",
        stops: [
          { position: 0, color: "FFFFFF" },
          { position: 100000, color: "4472C4" },
        ],
      },
    };
    const result = roundTrip(opts);
    const fill = result.fill! as { type: string };

    expect(fill.type).toBe("gradient");
  });

  it("round-trips patternFill", () => {
    const opts: BackgroundDescriptorOptions = {
      fill: {
        type: "pattern",
        pattern: "diagCross",
        foregroundColor: "FF0000",
        backgroundColor: "FFFFFF",
      },
    };
    const result = roundTrip(opts);
    const fill = result.fill! as { type: string };

    expect(fill.type).toBe("pattern");
  });
});
