import { parse as parseXml } from "@office-open/xml";
import { describe, it, expect } from "vite-plus/test";

import { stringify, parse } from "../descriptor";
import { themeDesc } from "./theme-descriptors";
import type { ThemeOptions } from "./theme-options";

function roundTrip(opts: ThemeOptions): ThemeOptions {
  const xml = stringify(themeDesc, opts, {} as any);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return parse(themeDesc, el, {} as any);
}

describe("themeDesc", () => {
  it("round-trips theme name", () => {
    const opts: ThemeOptions = { name: "Office Theme" };
    const result = roundTrip(opts);
    expect(result.name).toBe("Office Theme");
  });

  it("round-trips color scheme", () => {
    const opts: ThemeOptions = {
      colors: {
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
    expect(result.colors?.dark1).toBe("000000");
    expect(result.colors?.accent1).toBe("2563EB");
    expect(result.colors?.hyperlink).toBe("2563EB");
  });

  it("round-trips font scheme", () => {
    const opts: ThemeOptions = {
      fonts: {
        majorFont: "Calibri Light",
        minorFont: "Calibri",
      },
    };
    const result = roundTrip(opts);
    expect(result.fonts?.majorFont).toBe("Calibri Light");
    expect(result.fonts?.minorFont).toBe("Calibri");
  });

  it("round-trips full theme", () => {
    const opts: ThemeOptions = {
      name: "Custom Theme",
      colors: { accent1: "FF0000", dark1: "000000" },
      fonts: { majorFont: "Arial", minorFont: "Verdana" },
    };
    const result = roundTrip(opts);
    expect(result.name).toBe("Custom Theme");
    expect(result.colors?.accent1).toBe("FF0000");
    expect(result.fonts?.majorFont).toBe("Arial");
  });
});
