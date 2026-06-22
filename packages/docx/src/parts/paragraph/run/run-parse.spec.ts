import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { stringifyRunProperties } from "../stringify";
import type { RunPropertiesOptions } from "./properties";
import { parseRun, parseRunProperties } from "./run-parse";

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

  it("round-trips characterSpacing with UniversalMeasure (mm)", () => {
    const result = roundTrip({ characterSpacing: "0.5mm" });
    expect(result.characterSpacing).toBe("0.5mm");
  });

  // b/bCs, i/iCs, sz/szCs are independent toggle/measure properties (Latin vs
  // complex script, per ISO/IEC 29500). stringify must not auto-pair them
  // (bold → b+bCs would inflate), so round-trip stays field-faithful:
  // a source <w:b/> round-trips as <w:b/>, not <w:b/><w:bCs/>.
  it("does not auto-pair bCs when only bold is set", () => {
    const rPr = stringifyRunProperties({ bold: true })!;
    expect(rPr).toContain("<w:b");
    expect(rPr).not.toContain("bCs");
  });

  it("emits bCs only when boldComplexScript is explicitly set", () => {
    const rPr = stringifyRunProperties({ bold: true, boldComplexScript: true })!;
    expect(rPr).toContain("<w:b");
    expect(rPr).toContain("<w:bCs");
  });

  it("round-trips bold without inflating boldComplexScript", () => {
    const result = roundTrip({ bold: true });
    expect(result.bold).toBe(true);
    expect(result.boldComplexScript).toBeUndefined();
  });

  it("does not auto-pair iCs when only italic is set", () => {
    const rPr = stringifyRunProperties({ italic: true })!;
    expect(rPr).toContain("<w:i");
    expect(rPr).not.toContain("iCs");
  });

  it("does not auto-pair szCs when only size is set", () => {
    const rPr = stringifyRunProperties({ size: 24 })!;
    expect(rPr).toContain("<w:sz ");
    expect(rPr).not.toContain("szCs");
  });

  it("emits szCs when sizeComplexScript is an explicit number", () => {
    const rPr = stringifyRunProperties({ size: 24, sizeComplexScript: 20 })!;
    expect(rPr).toContain("<w:sz ");
    expect(rPr).toContain("<w:szCs");
  });
});

describe("parseRun rsid attributes", () => {
  it("reads w:rsidR/w:rsidRPr/w:rsidDel (hex verbatim, leading zeros kept)", () => {
    const doc = parseXml(
      `<w:r ${W_NS} w:rsidR="00992297" w:rsidRPr="00112233" w:rsidDel="AABBCCDD"><w:t>hi</w:t></w:r>`,
    );
    // parseRun does not use its context for rsid reads.
    const parsed = parseRun(doc.elements![0], {} as never);
    expect(parsed.rsid).toBe("00992297");
    expect(parsed.runPropertiesRsid).toBe("00112233");
    expect(parsed.deletionRsid).toBe("AABBCCDD");
  });
});
