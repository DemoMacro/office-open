import { parse as parseXml } from "@office-open/xml";
import { describe, it, expect } from "vite-plus/test";

import { stringify, parse } from "../../descriptor";
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

function roundTrip<T>(desc: any, opts: T): T {
  const xml = stringify(desc, opts, {} as any);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return parse(desc, el, {} as any);
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
      transforms: { tint: 50000, alpha: 80000 },
    };
    const result = roundTrip(rgbColorDesc, opts);
    expect(result.value).toBe("4472C4");
    expect(result.transforms).toBeDefined();
    expect(result.transforms!.tint).toBe(50000);
    expect(result.transforms!.alpha).toBe(80000);
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
      transforms: { shade: 25000 },
    };
    const result = roundTrip(schemeColorDesc, opts);
    expect(result.value).toBe("accent2");
    expect(result.transforms!.shade).toBe(25000);
  });
});

describe("hslColorDesc", () => {
  it("round-trips HSL color", () => {
    const opts: HslColorOptions = { hue: 120000, saturation: 100000, luminance: 50000 };
    const result = roundTrip(hslColorDesc, opts);
    expect(result.hue).toBe(120000);
    expect(result.saturation).toBe(100000);
    expect(result.luminance).toBe(50000);
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
    const opts: ScRgbColorOptions = { r: "0.5", g: "0.3", b: "0.8" };
    const result = roundTrip(scRgbColorDesc, opts);
    expect(result.r).toBe("0.5");
    expect(result.g).toBe("0.3");
    expect(result.b).toBe("0.8");
  });
});

describe("solidFillDesc", () => {
  it("round-trips solidFill with RGB color", () => {
    const opts: SolidFillOptions = { value: "FF0000" };
    const result = roundTrip(solidFillDesc, opts);
    expect(result).toEqual({ value: "FF0000" });
  });

  it("round-trips solidFill with scheme color", () => {
    const opts: SolidFillOptions = { value: "accent1", transforms: { tint: 50000 } };
    const result = roundTrip(solidFillDesc, opts);
    expect(result.value).toBe("accent1");
    expect(result.transforms!.tint).toBe(50000);
  });

  it("round-trips solidFill with HSL color", () => {
    const opts: SolidFillOptions = { hue: 240000, saturation: 80000, luminance: 60000 };
    const result = roundTrip(solidFillDesc, opts);
    expect(result.hue).toBe(240000);
    expect(result.saturation).toBe(80000);
    expect(result.luminance).toBe(60000);
  });

  it("round-trips solidFill with system color", () => {
    const opts: SolidFillOptions = { value: "windowText", lastClr: "000000" };
    const result = roundTrip(solidFillDesc, opts);
    expect(result.value).toBe("windowText");
    expect(result.lastClr).toBe("000000");
  });
});
