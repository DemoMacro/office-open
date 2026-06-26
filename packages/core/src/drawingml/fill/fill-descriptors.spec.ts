import { parse as parseXml } from "@office-open/xml";
import { describe, it, expect } from "vite-plus/test";

import { stringify, parse, type ReadContext, type WriteContext } from "../../descriptor";
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

  it("round-trips path gradient with fillToRectangle", () => {
    const opts: GradientFillOptions = {
      stops: [
        { position: 0, color: { value: "FFFFFF" } },
        { position: 100000, color: { value: "4472C4" } },
      ],
      shade: { path: "circle", fillToRectangle: { left: "50000", top: "50000" } },
    };
    const result = roundTripGradient(opts);
    expect(result.shade).toBeDefined();
    if (result.shade && "path" in result.shade) {
      expect(result.shade.path).toBe("circle");
      expect(result.shade.fillToRectangle).toBeDefined();
    }
  });

  it("round-trips gradient stop with non-rgb color (HSL)", () => {
    // a:gs holds EG_ColorChoice directly; readDirectColor must handle all six
    // color kinds, not just srgbClr/schemeClr (otherwise HSL/sys/prst/scRgb drop).
    const opts: GradientFillOptions = {
      stops: [
        { position: 0, color: { hue: 120000, saturation: 100000, luminance: 50000 } },
        { position: 100000, color: { value: "4472C4" } },
      ],
    };
    const result = roundTripGradient(opts);
    expect(result.stops).toHaveLength(2);
    expect(result.stops[0].color).toEqual({
      hue: 120000,
      saturation: 100000,
      luminance: 50000,
    });
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

  it("round-trips pattern with non-rgb foreground color (HSL)", () => {
    const opts: PatternFillOptions = {
      pattern: "cross",
      foregroundColor: { hue: 0, saturation: 100000, luminance: 50000 },
    };
    const result = roundTripPattern(opts);
    expect(result.foregroundColor).toEqual({ hue: 0, saturation: 100000, luminance: 50000 });
  });
});

describe("fillDesc blip fill (parse)", () => {
  // Blip fills can't round-trip through fillDesc.stringify — image media is
  // registered by the format-package compiler (not the core descriptor), so
  // stringify returns undefined for `type: "blip"`. These tests parse a
  // hand-built a:blipFill and assert the read-context media bridge resolves
  // r:embed → binary data + image type.
  const mockReadCtx = (overrides: Partial<ReadContext> = {}): ReadContext =>
    ({
      resolveRelationship: (rId: string) => (rId === "rId1" ? "media/image1.png" : undefined),
      getRaw: (path: string) =>
        path === "media/image1.png" ? new Uint8Array([1, 2, 3]) : undefined,
      getPart: () => undefined,
      ...overrides,
    }) as unknown as ReadContext;

  it("parses a:blipFill, bridging r:embed to binary media", () => {
    const xml = `<a:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`;
    const el = parseXml(xml).elements![0];
    const result = parse(fillDesc, el, mockReadCtx()) as Record<string, unknown>;
    expect(result.type).toBe("blip");
    expect(result.imageType).toBe("png");
    expect(Array.from(result.data as Uint8Array)).toEqual([1, 2, 3]);
  });

  it("preserves blip fill structure (dpi, rotWithShape)", () => {
    const xml = `<a:blipFill dpi="150" rotWithShape="1"><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`;
    const el = parseXml(xml).elements![0];
    const result = parse(fillDesc, el, mockReadCtx()) as Record<string, unknown>;
    expect(result.type).toBe("blip");
    expect(result.dpi).toBe(150);
    expect(result.rotWithShape).toBe(true);
  });

  it("infers image type from path extension (jpg)", () => {
    const xml = `<a:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`;
    const el = parseXml(xml).elements![0];
    const ctx = mockReadCtx({
      resolveRelationship: (rId) => (rId === "rId2" ? "media/photo.jpeg" : undefined),
      getRaw: (path) => (path === "media/photo.jpeg" ? new Uint8Array([9]) : undefined),
    });
    const result = parse(fillDesc, el, ctx) as Record<string, unknown>;
    expect(result.imageType).toBe("jpg");
  });

  it("falls back to none when media cannot be resolved", () => {
    const xml = `<a:blipFill><a:blip r:embed="missing"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`;
    const el = parseXml(xml).elements![0];
    const result = parse(fillDesc, el, mockReadCtx()) as Record<string, unknown>;
    expect(result.type).toBe("none");
  });
});

describe("fillDesc blip fill (stringify)", () => {
  // Blip stringify registers image media via ctx.addMedia and emits a:blipFill
  // with the returned {fileName} placeholder; the format-package compiler
  // replaces the placeholder with a relationship rId at pack time.
  const mockWriteCtx = (placeholder = "{image1.png}"): WriteContext =>
    ({
      addRelationship: () => "rId1",
      addMedia: (_data: Uint8Array, _type: string) => placeholder,
    }) as unknown as WriteContext;

  it("registers media and emits a:blipFill with the embed placeholder", () => {
    const opts: FillOptions = { type: "blip", data: new Uint8Array([1, 2, 3]), imageType: "png" };
    const xml = stringify(fillDesc, opts, mockWriteCtx());
    expect(xml).toContain("<a:blipFill");
    expect(xml).toContain('r:embed="{image1.png}"');
    expect(xml).toContain("<a:stretch><a:fillRect/></a:stretch>");
  });

  it("emits source rectangle and blip effects when provided", () => {
    const opts: FillOptions = {
      type: "blip",
      data: new Uint8Array([1]),
      imageType: "png",
      blipEffects: { grayscale: true },
      sourceRectangle: { left: 10000, top: 20000, right: 30000, bottom: 40000 },
    };
    const xml = stringify(fillDesc, opts, mockWriteCtx());
    expect(xml).toContain("<a:srcRect");
    expect(xml).toContain("a:grayscl");
  });
});
