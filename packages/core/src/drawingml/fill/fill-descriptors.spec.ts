import { parse as parseXml } from "@office-open/xml";
import { describe, it, expect } from "vite-plus/test";

import { stringify, parse } from "../../descriptor";
import { fillDesc, gradientFillDesc, patternFillDesc } from "./fill-descriptors";
import type { FillOptions } from "./fill-options";
import type { GradientFillOptions } from "./gradient-fill";
import type { PatternFillOptions } from "./pattern-fill";

function roundTripFill(opts: FillOptions): FillOptions {
  const xml = stringify(fillDesc, opts, {} as any);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return parse(fillDesc, el, {} as any);
}

function roundTripFillAsRecord(opts: FillOptions): Record<string, unknown> {
  return roundTripFill(opts) as unknown as Record<string, unknown>;
}

function roundTripGradient(opts: GradientFillOptions): GradientFillOptions {
  const xml = stringify(gradientFillDesc, opts, {} as any);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return parse(gradientFillDesc, el, {} as any);
}

function roundTripPattern(opts: PatternFillOptions): PatternFillOptions {
  const xml = stringify(patternFillDesc, opts, {} as any);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return parse(patternFillDesc, el, {} as any);
}

describe("fillDesc", () => {
  it("round-trips solid fill (string shorthand)", () => {
    const result = roundTripFill("FF0000");
    expect(result).toEqual({ type: "solid", color: { value: "FF0000" } });
  });

  it("round-trips solid fill (object)", () => {
    const opts: FillOptions = { type: "solid", color: { value: "4472C4" } };
    const result = roundTripFill(opts);
    expect(result).toEqual({ type: "solid", color: { value: "4472C4" } });
  });

  it("round-trips none fill", () => {
    const opts: FillOptions = { type: "none" };
    const result = roundTripFill(opts);
    expect(result).toEqual({ type: "none" });
  });

  it("round-trips group fill", () => {
    const opts: FillOptions = { type: "group" };
    const result = roundTripFill(opts);
    expect(result).toEqual({ type: "group" });
  });

  it("round-trips gradient fill (core API variant)", () => {
    const opts: FillOptions = {
      type: "gradient",
      options: {
        stops: [
          { position: 0, color: { value: "FF0000" } },
          { position: 100000, color: { value: "0000FF" } },
        ],
        shade: { angle: 5400000, scaled: true },
      },
    };
    const result = roundTripFillAsRecord(opts);
    expect(result.type).toBe("gradient");
    if ("options" in result && result.options) {
      const options = result.options as Record<string, unknown>;
      const stops = options.stops as Record<string, unknown>[];
      expect(stops).toHaveLength(2);
      expect(stops[0].position).toBe(0);
      expect(stops[1].position).toBe(100000);
    } else {
      expect.fail("expected options in gradient fill result");
    }
  });

  it("round-trips pattern fill", () => {
    const opts: FillOptions = {
      type: "pattern",
      pattern: "cross",
      foregroundColor: { value: "FF0000" },
      backgroundColor: { value: "FFFFFF" },
    };
    const result = roundTripFillAsRecord(opts);
    expect(result.type).toBe("pattern");
    if (result.type === "pattern") {
      expect(result.foregroundColor).toEqual({ value: "FF0000" });
      expect(result.backgroundColor).toEqual({ value: "FFFFFF" });
    }
  });
});

describe("gradientFillDesc", () => {
  it("round-trips linear gradient", () => {
    const opts: GradientFillOptions = {
      stops: [
        { position: 0, color: { value: "4472C4" } },
        { position: 50000, color: { value: "ED7D31" } },
        { position: 100000, color: { value: "A5A5A5" } },
      ],
      shade: { angle: 5400000, scaled: false },
      flip: "x",
      rotateWithShape: true,
    };
    const result = roundTripGradient(opts);
    expect(result.stops).toHaveLength(3);
    expect(result.stops[0].position).toBe(0);
    expect(result.stops[1].position).toBe(50000);
    expect(result.stops[2].position).toBe(100000);
    expect(result.shade).toBeDefined();
    if (result.shade && "angle" in result.shade) {
      expect(result.shade.angle).toBe(5400000);
      expect(result.shade.scaled).toBe(false);
    }
    expect(result.flip).toBe("x");
    expect(result.rotateWithShape).toBe(true);
  });

  it("round-trips path gradient with fillToRect", () => {
    const opts: GradientFillOptions = {
      stops: [
        { position: 0, color: { value: "FFFFFF" } },
        { position: 100000, color: { value: "4472C4" } },
      ],
      shade: { path: "circle", fillToRect: { left: "50000", top: "50000" } },
    };
    const result = roundTripGradient(opts);
    expect(result.shade).toBeDefined();
    if (result.shade && "path" in result.shade) {
      expect(result.shade.path).toBe("circle");
      expect(result.shade.fillToRect).toBeDefined();
    }
  });
});

describe("patternFillDesc", () => {
  it("round-trips pattern with colors", () => {
    const opts: PatternFillOptions = {
      pattern: "cross",
      foregroundColor: { value: "333333" },
      backgroundColor: { value: "CCCCCC" },
    };
    const result = roundTripPattern(opts);
    expect(result.foregroundColor).toEqual({ value: "333333" });
    expect(result.backgroundColor).toEqual({ value: "CCCCCC" });
  });
});
