import { Formatter } from "@export/formatter";
import { toElement } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { Paragraph } from "../paragraph";
import { EmphasisMarkType } from "./emphasis-mark";
import { SymbolRun } from "./symbol-run";
import { UnderlineType } from "./underline";

describe("SymbolRun", () => {
  let run: SymbolRun;

  describe("#constructor()", () => {
    it("should create symbol run from text input", () => {
      run = new SymbolRun("F071");
      const f = new Formatter().format(run);
      expect(f).to.deep.equal({
        "w:r": [{ "w:sym": { _attr: { "w:char": "F071", "w:font": "Wingdings" } } }],
      });
    });

    it("should create symbol run from object input with just 'char' specified", () => {
      run = new SymbolRun({ char: "F071" });
      const f = new Formatter().format(run);
      expect(f).to.deep.equal({
        "w:r": [{ "w:sym": { _attr: { "w:char": "F071", "w:font": "Wingdings" } } }],
      });
    });

    it("should create symbol run from object input with just 'char' specified", () => {
      run = new SymbolRun({ char: "F071", symbolfont: "Arial" });
      const f = new Formatter().format(run);
      expect(f).to.deep.equal({
        "w:r": [{ "w:sym": { _attr: { "w:char": "F071", "w:font": "Arial" } } }],
      });
    });

    it("should add other standard run properties", () => {
      run = new SymbolRun({
        bold: true,
        char: "F071",
        color: "00FF00",
        emphasisMark: {
          type: EmphasisMarkType.DOT,
        },
        highlight: "yellow",
        italics: true,
        size: 40,
        symbolfont: "Arial",
        underline: {
          color: "ff0000",
          type: UnderlineType.DOUBLE,
        },
      });

      const f = new Formatter().format(run);
      expect(f).to.deep.equal({
        "w:r": [
          {
            "w:rPr": [
              { "w:b": {} },
              { "w:bCs": {} },
              { "w:i": {} },
              { "w:iCs": {} },
              { "w:color": { _attr: { "w:val": "00FF00" } } },
              { "w:sz": { _attr: { "w:val": 40 } } },
              { "w:szCs": { _attr: { "w:val": 40 } } },
              { "w:highlight": { _attr: { "w:val": "yellow" } } },
              { "w:u": { _attr: { "w:color": "ff0000", "w:val": "double" } } },
              { "w:em": { _attr: { "w:val": "dot" } } },
            ],
          },
          {
            "w:sym": { _attr: { "w:char": "F071", "w:font": "Arial" } },
          },
        ],
      });
    });
  });
});

// ── Parse round-trip tests ──────────────────────────────────────────────────

describe("SymbolRun parse round-trip", () => {
  it("should round-trip symbolRun via Paragraph JSON API", () => {
    const paragraph = new Paragraph({
      children: [{ symbolRun: { char: "F071", symbolfont: "Wingdings" } }],
    });
    const tree = new Formatter().format(paragraph);
    const root = toElement(tree);
    // Find w:r containing w:sym
    const symRun = root.elements?.find((e) => e.elements?.some((c) => c.name === "w:sym"));
    expect(symRun).toBeDefined();
    const symEl = symRun!.elements!.find((c) => c.name === "w:sym");
    // toElement converts _attr to attributes
    expect(symEl?.attributes?.["w:char"]).toBe("F071");
    expect(symEl?.attributes?.["w:font"]).toBe("Wingdings");
  });

  it("should round-trip symbolRun with formatting via Paragraph JSON API", () => {
    const paragraph = new Paragraph({
      children: [
        { symbolRun: { char: "F021", symbolfont: "Wingdings", bold: true, color: "FF0000" } },
      ],
    });
    const tree = new Formatter().format(paragraph);
    const root = toElement(tree);
    const symRun = root.elements?.find((e) => e.elements?.some((c) => c.name === "w:sym"));
    expect(symRun).toBeDefined();
    // Should have run properties
    const rPr = symRun!.elements!.find((c) => c.name === "w:rPr");
    expect(rPr).toBeDefined();
  });
});
