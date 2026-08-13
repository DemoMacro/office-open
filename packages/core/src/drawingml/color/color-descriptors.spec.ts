import { parse as parseXml } from "@office-open/xml";
import { describe, it, expect } from "vite-plus/test";

import {
  stringify,
  parse,
  type CustomDescriptor,
  type ReadContext,
  type WriteContext,
} from "../../descriptor";
import {
  rgbColorDesc,
  schemeColorDesc,
  solidFillDesc,
  hslColorDesc,
  systemColorDesc,
  presetColorDesc,
  scRgbColorDesc,
} from "./color-descriptors";
import type { HslColorOptions } from "./hsl-color";
import type { PresetColorOptions } from "./preset-color";
import type { RgbColorOptions } from "./rgb-color";
import type { ScRgbColorOptions } from "./sc-rgb-color";
import type { SchemeColorOptions } from "./scheme-color";
import type { SolidFillOptions } from "./solid-fill";
import type { SystemColorOptions } from "./system-color";

function roundTrip<T>(desc: CustomDescriptor<T>, opts: T): T {
  const xml = stringify(desc, opts, {} as WriteContext);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return parse(desc, el, {} as ReadContext);
}

describe("rgbColorDesc", () => {
  it("round-trips basic RGB color", () => {
    const opts: RgbColorOptions = { value: "FF0000" };
    const result = roundTrip(rgbColorDesc, opts);
    expect(result.value).toBe("FF0000");
    expect(result.transforms).toBeUndefined();
  });

  it("round-trips RGB color with transforms", () => {
    const opts: RgbColorOptions = {
      value: "4472C4",
      transforms: { tint: 50, alpha: 80 },
    };
    const result = roundTrip(rgbColorDesc, opts);
    expect(result.value).toBe("4472C4");
    expect(result.transforms).toBeDefined();
    expect(result.transforms!.tint).toBe(50);
    expect(result.transforms!.alpha).toBe(80);
  });
});

describe("schemeColorDesc", () => {
  it("round-trips basic scheme color", () => {
    const opts: SchemeColorOptions = { value: "accent1" };
    const result = roundTrip(schemeColorDesc, opts);
    expect(result.value).toBe("accent1");
  });

  it("round-trips scheme color with transforms", () => {
    const opts: SchemeColorOptions = {
      value: "accent2",
      transforms: { shade: 25 },
    };
    const result = roundTrip(schemeColorDesc, opts);
    expect(result.value).toBe("accent2");
    expect(result.transforms!.shade).toBe(25);
  });
});

describe("hslColorDesc", () => {
  it("round-trips HSL color", () => {
    const opts: HslColorOptions = { hue: 120000, saturation: 100, luminance: 50 };
    const result = roundTrip(hslColorDesc, opts);
    expect(result.hue).toBe(120000);
    expect(result.saturation).toBe(100);
    expect(result.luminance).toBe(50);
  });
});

describe("systemColorDesc", () => {
  it("round-trips system color", () => {
    const opts: SystemColorOptions = { value: "windowText", lastClr: "000000" };
    const result = roundTrip(systemColorDesc, opts);
    expect(result.value).toBe("windowText");
    expect(result.lastClr).toBe("000000");
  });
});

describe("presetColorDesc", () => {
  it("round-trips preset color", () => {
    const opts: PresetColorOptions = { value: "blue" };
    const result = roundTrip(presetColorDesc, opts);
    expect(result.value).toBe("blue");
  });
});

describe("scRgbColorDesc", () => {
  it("round-trips scRGB color", () => {
    const opts: ScRgbColorOptions = { r: 50, g: 30, b: 80 };
    const result = roundTrip(scRgbColorDesc, opts);
    expect(result.r).toBe(50);
    expect(result.g).toBe(30);
    expect(result.b).toBe(80);
  });
});

describe("solidFillDesc", () => {
  it("round-trips solidFill with RGB color", () => {
    const opts: SolidFillOptions = { value: "FF0000" };
    const result = roundTrip(solidFillDesc, opts);
    expect(result).toEqual({ value: "FF0000" });
  });

  it("round-trips solidFill with scheme color", () => {
    const opts: SolidFillOptions = { value: "accent1", transforms: { tint: 50 } };
    const result = roundTrip(solidFillDesc, opts) as SchemeColorOptions;
    expect(result.value).toBe("accent1");
    expect(result.transforms!.tint).toBe(50);
  });

  it("round-trips solidFill with HSL color", () => {
    const opts: SolidFillOptions = { hue: 240000, saturation: 80, luminance: 60 };
    const result = roundTrip(solidFillDesc, opts) as HslColorOptions;
    expect(result.hue).toBe(240000);
    expect(result.saturation).toBe(80);
    expect(result.luminance).toBe(60);
  });

  it("round-trips solidFill with system color", () => {
    const opts: SolidFillOptions = { value: "windowText", lastClr: "000000" };
    const result = roundTrip(solidFillDesc, opts) as SystemColorOptions;
    expect(result.value).toBe("windowText");
    expect(result.lastClr).toBe("000000");
  });
});
