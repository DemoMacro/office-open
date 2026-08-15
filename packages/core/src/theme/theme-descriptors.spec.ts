import { parse as parseXml } from "@office-open/xml";
import { describe, it, expect } from "vite-plus/test";

import { stringify, parse, type ReadContext, type WriteContext } from "../descriptor";
import { buildThemeXml } from "./build-theme-xml";
import { themeDesc } from "./theme-descriptors";
import type { ThemeOptions, ThemeOverrideOptions } from "./theme-options";
import { themeManagerDesc, themeOverrideDesc } from "./theme-override";

function roundTrip(opts: ThemeOptions): ThemeOptions {
  const xml = stringify(themeDesc, opts, {} as WriteContext);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return parse(themeDesc, el, {} as ReadContext);
}

describe("themeDesc", () => {
  it("round-trips theme name", () => {
    const result = roundTrip({ name: "Office Theme" });
    expect(result.name).toBe("Office Theme");
  });

  it("round-trips color scheme", () => {
    const opts: ThemeOptions = {
      colorScheme: {
        dark1: "000000",
        light1: "FFFFFF",
        dark2: "1F2937",
        light2: "F9FAFB",
        accent1: "2563EB",
        accent2: "7C3AED",
        accent3: "DB2777",
        accent4: "DC2626",
        accent5: "EA580C",
        accent6: "CA8A04",
        hyperlink: "2563EB",
        followedHyperlink: "7C3AED",
      },
    };
    const result = roundTrip(opts);
    expect(result.colorScheme?.dark1).toBe("000000");
    expect(result.colorScheme?.accent1).toBe("2563EB");
    expect(result.colorScheme?.hyperlink).toBe("2563EB");
  });

  it("round-trips extra color-scheme mappings with full-word API values", () => {
    const opts: ThemeOptions = {
      extraColorSchemes: [
        {
          colorScheme: { accent1: "112233" },
          colorMapping: { background1: "dark1", followedHyperlink: "hyperlink" },
        },
      ],
    };
    const xml = stringify(themeDesc, opts, {} as WriteContext)!;
    expect(xml).toContain('<a:clrMap bg1="dk1"');
    expect(xml).toContain('folHlink="hlink"');

    const result = roundTrip(opts);
    expect(result.extraColorSchemes?.[0]?.colorMapping).toEqual({
      background1: "dark1",
      text1: "dark1",
      background2: "light2",
      text2: "dark2",
      accent1: "accent1",
      accent2: "accent2",
      accent3: "accent3",
      accent4: "accent4",
      accent5: "accent5",
      accent6: "accent6",
      hyperlink: "hyperlink",
      followedHyperlink: "hyperlink",
    });
  });

  it("round-trips font scheme with font collections", () => {
    const opts: ThemeOptions = {
      fontScheme: {
        majorFont: { latin: { typeface: "Calibri Light" } },
        minorFont: { latin: { typeface: "Calibri" } },
      },
    };
    const result = roundTrip(opts);
    expect(result.fontScheme?.majorFont?.latin?.typeface).toBe("Calibri Light");
    expect(result.fontScheme?.minorFont?.latin?.typeface).toBe("Calibri");
  });

  it("round-trips east-asian and complex-script fonts", () => {
    const opts: ThemeOptions = {
      fontScheme: {
        majorFont: {
          latin: { typeface: "Calibri Light" },
          eastAsian: { typeface: "微软雅黑" },
          complexScript: { typeface: "Times New Roman" },
        },
      },
    };
    const result = roundTrip(opts);
    expect(result.fontScheme?.majorFont?.eastAsian?.typeface).toBe("微软雅黑");
    expect(result.fontScheme?.majorFont?.complexScript?.typeface).toBe("Times New Roman");
  });

  it("parses default theme with full format scheme", () => {
    const freshXml = buildThemeXml();
    const doc = parseXml(freshXml);
    const el = doc.elements?.[0];
    if (!el) throw new Error("no root element");
    const result = parse(themeDesc, el, {} as ReadContext);
    expect(result.colorScheme?.accent1).toBe("4472C4");
    expect(result.fontScheme?.majorFont?.latin?.typeface).toBe("Calibri Light");
    expect(result.formatScheme?.fillStyles).toHaveLength(3);
    expect(result.formatScheme?.lineStyles).toHaveLength(3);
    expect(result.formatScheme?.effectStyles).toHaveLength(3);
    expect(result.formatScheme?.backgroundFillStyles).toHaveLength(3);
  });

  it("round-trips full theme", () => {
    const opts: ThemeOptions = {
      name: "Custom Theme",
      colorScheme: { accent1: "FF0000", dark1: "000000" },
      fontScheme: {
        majorFont: { latin: { typeface: "Arial" } },
        minorFont: { latin: { typeface: "Verdana" } },
      },
    };
    const result = roundTrip(opts);
    expect(result.name).toBe("Custom Theme");
    expect(result.colorScheme?.accent1).toBe("FF0000");
    expect(result.fontScheme?.majorFont?.latin?.typeface).toBe("Arial");
  });
});

describe("custom colors", () => {
  it("round-trips custClrLst entries", () => {
    const opts: ThemeOptions = {
      customColors: [
        { name: "Brand Red", color: { value: "FF0000" } },
        { color: { value: "00FF00" } },
      ],
    };
    const xml = stringify(themeDesc, opts, {} as WriteContext)!;
    expect(xml).toContain('<a:custClr name="Brand Red">');

    const result = roundTrip(opts);
    expect(result.customColors).toHaveLength(2);
    expect(result.customColors?.[0]).toMatchObject({ name: "Brand Red" });
    expect(result.customColors?.[1]?.name).toBeUndefined();
  });

  it("omits custClrLst when no custom colors set", () => {
    const xml = stringify(themeDesc, {}, {} as WriteContext)!;
    expect(xml).not.toContain("custClrLst");
  });
});

describe("themeOverrideDesc", () => {
  it("round-trips the base-styles subset", () => {
    const opts: ThemeOverrideOptions = {
      colorScheme: { accent1: "123456", dark1: "000000" },
      fontScheme: { majorFont: { latin: { typeface: "Arial" } } },
    };
    const xml = stringify(themeOverrideDesc, opts, {} as WriteContext)!;
    expect(xml).toContain("<a:themeOverride ");
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("no root element");
    const result = parse(themeOverrideDesc, el, {} as ReadContext);
    expect(result.colorScheme?.accent1).toBe("123456");
    expect(result.fontScheme?.majorFont?.latin?.typeface).toBe("Arial");
    expect(result.formatScheme).toBeUndefined();
  });
});

describe("themeManagerDesc", () => {
  it("serializes the empty legacy element and parses back empty", () => {
    const xml = stringify(themeManagerDesc, {}, {} as WriteContext)!;
    expect(xml).toBe(
      '<a:themeManager xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>',
    );
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("no root element");
    expect(parse(themeManagerDesc, el, {} as ReadContext)).toEqual({});
  });
});
