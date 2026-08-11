import { parse as parseXml } from "@office-open/xml";
import { describe, it, expect } from "vite-plus/test";

import { stringify, parse, type ReadContext, type WriteContext } from "../descriptor";
import { buildThemeXml } from "./build-theme-xml";
import { themeDesc } from "./theme-descriptors";
import type { ThemeOptions } from "./theme-options";

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
