import { parse as parseXml } from "@office-open/xml";
import { describe, it, expect } from "vite-plus/test";

import { stringify, parse } from "../../descriptor";
import { effectListDesc } from "./effect-descriptors";
import type { EffectListOptions } from "./effect-list";

function roundTrip(opts: EffectListOptions): EffectListOptions {
  const xml = stringify(effectListDesc, opts, {} as any);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return parse(effectListDesc, el, {} as any);
}

describe("effectListDesc", () => {
  it("round-trips blur", () => {
    const opts: EffectListOptions = {
      blur: { radius: 50000, grow: false },
    };
    const result = roundTrip(opts);
    expect(result.blur).toBeDefined();
    expect(result.blur!.radius).toBe(50000);
    expect(result.blur!.grow).toBe(false);
  });

  it("round-trips blur with grow true (default)", () => {
    const opts: EffectListOptions = {
      blur: { radius: 25000 },
    };
    const result = roundTrip(opts);
    expect(result.blur!.radius).toBe(25000);
    // grow defaults to true in XSD; not emitted when true, so parse gives undefined
  });

  it("round-trips fillOverlay", () => {
    const opts: EffectListOptions = {
      fillOverlay: { blend: "over" },
    };
    const result = roundTrip(opts);
    expect(result.fillOverlay).toBeDefined();
    expect(result.fillOverlay!.blend).toBe("over");
  });

  it("emits XSD tokens for blend (mult) and alignments (ctr), not full words", () => {
    const xml = stringify(
      effectListDesc,
      {
        fillOverlay: { blend: "multiply" },
        outerShadow: { alignment: "center", color: { value: "000000" } },
        reflection: { alignment: "center" },
      },
      {} as any,
    )!;
    expect(xml).toContain('blend="mult"');
    expect(xml).not.toContain('blend="multiply"');
    expect(xml).toContain('algn="ctr"');
    expect(xml).not.toContain('algn="center"');
  });

  it("round-trips glow with RGB color", () => {
    const opts: EffectListOptions = {
      glow: { radius: 40000, color: { value: "FF0000" } },
    };
    const result = roundTrip(opts);
    expect(result.glow).toBeDefined();
    expect(result.glow!.radius).toBe(40000);
    expect(result.glow!.color).toEqual({ value: "FF0000" });
  });

  it("round-trips innerShadow", () => {
    const opts: EffectListOptions = {
      innerShadow: {
        blurRadius: 50000,
        distance: 25000,
        direction: 2700000,
        color: { value: "000000" },
      },
    };
    const result = roundTrip(opts);
    expect(result.innerShadow).toBeDefined();
    expect(result.innerShadow!.blurRadius).toBe(50000);
    expect(result.innerShadow!.distance).toBe(25000);
    expect(result.innerShadow!.direction).toBe(2700000);
    expect(result.innerShadow!.color).toEqual({ value: "000000" });
  });

  it("round-trips outerShadow", () => {
    const opts: EffectListOptions = {
      outerShadow: {
        blurRadius: 76200,
        distance: 40000,
        direction: 5400000,
        scaleX: 100000,
        scaleY: 100000,
        skewX: 0,
        skewY: 0,
        alignment: "bottom",
        rotWithShape: false,
        color: { value: "333333" },
      },
    };
    const result = roundTrip(opts);
    expect(result.outerShadow).toBeDefined();
    expect(result.outerShadow!.blurRadius).toBe(76200);
    expect(result.outerShadow!.distance).toBe(40000);
    expect(result.outerShadow!.direction).toBe(5400000);
    expect(result.outerShadow!.scaleX).toBe(100000);
    expect(result.outerShadow!.scaleY).toBe(100000);
    expect(result.outerShadow!.skewX).toBe(0);
    expect(result.outerShadow!.skewY).toBe(0);
    expect(result.outerShadow!.alignment).toBe("bottom");
    expect(result.outerShadow!.rotWithShape).toBe(false);
    expect(result.outerShadow!.color).toEqual({ value: "333333" });
  });

  it("round-trips presetShadow", () => {
    const opts: EffectListOptions = {
      presetShadow: {
        preset: "shadow2",
        distance: 30000,
        direction: 2700000,
        color: { value: "666666" },
      },
    };
    const result = roundTrip(opts);
    expect(result.presetShadow).toBeDefined();
    expect(result.presetShadow!.preset).toBe("shadow2");
    expect(result.presetShadow!.distance).toBe(30000);
    expect(result.presetShadow!.direction).toBe(2700000);
    expect(result.presetShadow!.color).toEqual({ value: "666666" });
  });

  it("round-trips reflection with all attributes", () => {
    const opts: EffectListOptions = {
      reflection: {
        blurRadius: 6350,
        startAlpha: 50000,
        startPosition: 0,
        endAlpha: 0,
        endPosition: 50000,
        distance: 25400,
        direction: 5400000,
        fadeDirection: 5400000,
        scaleX: 100000,
        scaleY: 100000,
        skewX: 0,
        skewY: 0,
        alignment: "bottom",
        rotWithShape: false,
      },
    };
    const result = roundTrip(opts);
    expect(result.reflection).toBeDefined();
    const ref = result.reflection as Record<string, unknown>;
    expect(ref.blurRadius).toBe(6350);
    expect(ref.startAlpha).toBe(50000);
    expect(ref.startPosition).toBe(0);
    expect(ref.endAlpha).toBe(0);
    expect(ref.endPosition).toBe(50000);
    expect(ref.distance).toBe(25400);
    expect(ref.direction).toBe(5400000);
    expect(ref.fadeDirection).toBe(5400000);
    expect(ref.scaleX).toBe(100000);
    expect(ref.scaleY).toBe(100000);
    expect(ref.skewX).toBe(0);
    expect(ref.skewY).toBe(0);
    expect(ref.alignment).toBe("bottom");
    expect(ref.rotWithShape).toBe(false);
  });

  it("round-trips softEdge", () => {
    const opts: EffectListOptions = {
      softEdge: 25000,
    };
    const result = roundTrip(opts);
    expect(result.softEdge).toBe(25000);
  });

  it("round-trips multiple effects combined", () => {
    const opts: EffectListOptions = {
      blur: { radius: 10000 },
      glow: { radius: 20000, color: { value: "4472C4" } },
      softEdge: 15000,
    };
    const result = roundTrip(opts);
    expect(result.blur!.radius).toBe(10000);
    expect(result.glow!.radius).toBe(20000);
    expect(result.glow!.color).toEqual({ value: "4472C4" });
    expect(result.softEdge).toBe(15000);
  });
});
