import { checkOrder, diffTagSets, findFieldSpec } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { describe, it, expect } from "vite-plus/test";

import { parseParagraphProperties } from "../../body";
import type { DocxReadContext } from "../../context";
import type { ParagraphPropertiesOptions } from "./properties";
import { stringifyParagraphProperties } from "./stringify";

const spec = findFieldSpec("paragraph-properties")!;

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

// parseParagraphProperties only reaches ctx inside the w:numPr branch; the
// probes below carry no numbering, so a stub context is safe at runtime.
const stubCtx = {} as unknown as DocxReadContext;

function parsePPr(inner: string): Record<string, unknown> {
  const xml = `<w:pPr xmlns:w="${W}">${inner}</w:pPr>`;
  const doc = parseXml(xml) as Element;
  const root = doc.elements?.find((e) => e.type === "element");
  if (!root) throw new Error("parsePPr: produced no root element");
  return parseParagraphProperties(root, stubCtx);
}

describe("paragraph-properties field consistency", () => {
  it("round-trips textDirection, textboxTightWrap, divId, cnfStyle", () => {
    const opts = parsePPr(
      `<w:textDirection w:val="lr"/>` +
        `<w:textboxTightWrap w:val="none"/>` +
        `<w:divId w:val="1"/>` +
        `<w:cnfStyle w:val="000000010000"/>`,
    );
    expect(opts.textDirection).toBe("lr");
    expect(opts.textboxTightWrap).toBe("none");
    expect(opts.divId).toBe(1);
    // w:val="000000010000" — 12-char ST_Cnf, digit index 7 set → evenHBand.
    expect(opts.cnfStyle).toEqual({ evenHBand: true });
  });

  it("round-trips revision (w:pPrChange)", () => {
    const opts = parsePPr(
      `<w:pPrChange w:id="1" w:author="a" w:date="2024-01-01T00:00:00Z">` +
        `<w:pPr><w:jc w:val="left"/></w:pPr>` +
        `</w:pPrChange>`,
    );
    expect(opts.revision).toEqual({
      id: 1,
      author: "a",
      date: "2024-01-01T00:00:00Z",
      alignment: "left",
    });
  });

  it("keeps numPr partial forms distinct (ilvl+numId / numId-only / ilvl-only)", () => {
    const numberingCtx = {
      numberingCache: new Map([["0", parseXml(`<w:abstractNum xmlns:w="${W}"/>`) as Element]]),
      numIdCache: new Map([["1", "0"]]),
    } as unknown as DocxReadContext;

    const parseWith = (inner: string): Record<string, unknown> => {
      const xml = `<w:pPr xmlns:w="${W}">${inner}</w:pPr>`;
      const root = (parseXml(xml) as Element).elements?.find((e) => e.type === "element");
      if (!root) throw new Error("parseWith: produced no root element");
      return parseParagraphProperties(root, numberingCtx);
    };

    // Full form → numbering with pinned level.
    const full = parseWith(`<w:numPr><w:ilvl w:val="2"/><w:numId w:val="1"/></w:numPr>`);
    expect(full.numbering).toMatchObject({ reference: "list_1", level: 2 });

    // numId-only (no w:ilvl) → level stays undefined so stringify omits w:ilvl.
    const numIdOnly = parseWith(`<w:numPr><w:numId w:val="1"/></w:numPr>`);
    expect(numIdOnly.numbering).toMatchObject({ reference: "list_1" });
    expect((numIdOnly.numbering as { level?: number }).level).toBeUndefined();

    // ilvl-only (no w:numId) → inherits numbering from the style chain; no
    // bullet fallback (which would fabricate numId=1 + ListParagraph).
    const ilvlOnly = parseWith(`<w:numPr><w:ilvl w:val="1"/></w:numPr>`);
    expect(ilvlOnly.numbering).toBeUndefined();
    expect(ilvlOnly.bullet).toBeUndefined();
  });

  it("declared F3 parse-loss matches the live parse gap (regression guard)", () => {
    // If parseParagraphProperties gains a findChild for any of these, the
    // field set above must be updated too — this keeps FIELD_SPECS honest.
    const report = diffTagSets(spec);
    expect(report.f3ParseLoss).toEqual([]);
    expect(report.f1WriteLoss).toEqual([]);
    expect(report.f2WriteOnly).toEqual([]);
    expect(report.f5ParseOnly).toEqual([]);
  });

  it("F6 — child order matches the XSD pPr sequence", () => {
    // pPr is an XSD sequence (EG_pPrBaseOrder); out-of-order children make
    // Word reject the part. The sample exercises 14 distinct child elements,
    // enough to catch a transposition.
    const result = stringifyParagraphProperties(
      spec.sampleOptions as unknown as ParagraphPropertiesOptions,
    );
    const violations = checkOrder(result.xml!, spec.order!);
    expect(violations).toEqual([]);
  });
});
