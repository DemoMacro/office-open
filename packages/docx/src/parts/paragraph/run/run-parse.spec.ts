import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { stringifyRunProperties } from "../stringify";
import type { RunPropertiesOptions } from "./properties";
import { parseRunProperties } from "./run-parse";

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function roundTrip(opts: RunPropertiesOptions): RunPropertiesOptions {
  const rPr = stringifyRunProperties(opts)!;
  const doc = parseXml(`<w:r ${W_NS}>${rPr}</w:r>`);
  const r = doc.elements![0];
  const rPrEl = r.elements![0];
  return parseRunProperties(rPrEl);
}

describe("parseRunProperties round-trip", () => {
  it("round-trips color with themeColor/themeTint/themeShade", () => {
    const result = roundTrip({
      color: { val: "FF0000", themeColor: "text1", themeTint: "99", themeShade: "BF" },
    });
    expect(result.color).toEqual({
      val: "FF0000",
      themeColor: "text1",
      themeTint: "99",
      themeShade: "BF",
    });
  });

  it("round-trips color as plain string when no theme attributes", () => {
    const result = roundTrip({ color: "FF0000" });
    expect(result.color).toBe("FF0000");
  });

  it("round-trips eastAsianLayout", () => {
    const result = roundTrip({
      eastAsianLayout: {
        id: 1,
        combine: true,
        combineBrackets: "round",
        vert: true,
        vertCompress: false,
      },
    });
    expect(result.eastAsianLayout).toEqual({
      id: 1,
      combine: true,
      combineBrackets: "round",
      vert: true,
      vertCompress: false,
    });
  });

  it("round-trips contentPartRId", () => {
    const result = roundTrip({ contentPartRId: "rId42" });
    expect(result.contentPartRId).toBe("rId42");
  });

  it("round-trips rPrChange revision with inner rPr", () => {
    const result = roundTrip({
      bold: true,
      revision: { id: 1, author: "A", date: "2024-01-01T00:00:00Z", italic: true },
    });
    expect(result.revision).toBeDefined();
    const rev = result.revision as unknown as Record<string, unknown>;
    expect(rev.id).toBe(1);
    expect(rev.author).toBe("A");
    expect(rev.date).toBe("2024-01-01T00:00:00Z");
    expect(rev.italic).toBe(true);
  });
});
