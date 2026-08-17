import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { BackgroundOptions } from "../background";
import { backgroundDesc } from "./background";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: BackgroundOptions) {
  const xml = backgroundDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return backgroundDesc.parse(el, readCtx);
}

describe("backgroundDesc round-trip", () => {
  it("round-trips solidFill color", () => {
    const opts: BackgroundOptions = {
      fill: { type: "solid", color: "FF0000" },
    };
    const result = roundTrip(opts);
    const fill = result.fill! as { type: string; color: { value: string } };

    expect(fill.type).toBe("solid");
    expect(fill.color.value).toBe("FF0000");
  });

  it("round-trips noFill", () => {
    const opts: BackgroundOptions = {
      fill: { type: "none" },
    };
    const result = roundTrip(opts);
    const fill = result.fill! as { type: string };

    expect(fill.type).toBe("none");
  });

  it("round-trips shadeToTitle", () => {
    const opts: BackgroundOptions = {
      fill: { type: "solid", color: "4472C4" },
      shadeToTitle: true,
    };
    const result = roundTrip(opts);

    expect(result.shadeToTitle).toBe(true);
  });

  it("round-trips blackWhiteMode", () => {
    const opts: BackgroundOptions = {
      fill: { type: "solid", color: "000000" },
      blackWhiteMode: "gray",
    };
    const result = roundTrip(opts);

    expect(result.blackWhiteMode).toBe("gray");
  });

  it("emits unprefixed bwMode on p:bg", () => {
    // CT_Background declares @bwMode in no namespace, not p:bwMode.
    const xml = backgroundDesc.stringify(
      { fill: { type: "solid", color: "000000" }, blackWhiteMode: "gray" },
      {} as never,
    );
    expect(xml).toContain('<p:bg bwMode="gray">');
    expect(xml).not.toContain("p:bwMode");
  });

  it("round-trips a style matrix reference (p:bgRef)", () => {
    const opts: BackgroundOptions = {
      reference: { index: 1001, color: { value: "bg1" } },
    };
    const result = roundTrip(opts);

    expect(result.reference).toEqual({ index: 1001, color: { value: "bg1" } });
  });

  it("emits p:bgRef without a color child when color is omitted", () => {
    const xml = backgroundDesc.stringify({ reference: { index: 1002 } }, {} as never)!;
    expect(xml).toBe('<p:bg><p:bgRef idx="1002"></p:bgRef></p:bg>');
  });

  it("round-trips gradientFill", () => {
    const opts: BackgroundOptions = {
      fill: {
        type: "gradient",
        stops: [
          { position: 0, color: "FFFFFF" },
          { position: 100, color: "4472C4" },
        ],
      },
    };
    const result = roundTrip(opts);
    const fill = result.fill! as { type: string };

    expect(fill.type).toBe("gradient");
  });

  it("round-trips patternFill", () => {
    const opts: BackgroundOptions = {
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

  it("round-trips effects (a:effectLst outerShadow)", () => {
    const opts: BackgroundOptions = {
      fill: { type: "solid", color: "FF0000" },
      effects: {
        outerShadow: {
          color: { value: "000000" },
          blurRadius: 50000,
          distance: 30000,
          direction: 45,
        },
      },
    };
    const result = roundTrip(opts);
    const effects = result.effects!;
    const shadow = effects.outerShadow!;

    expect((shadow.color as { value: string }).value).toBe("000000");
    expect(shadow.blurRadius).toBe(50000);
    expect(shadow.distance).toBe(30000);
    expect(shadow.direction).toBe(45);
  });
});
