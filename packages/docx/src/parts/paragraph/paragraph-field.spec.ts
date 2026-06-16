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
  it("F3 — parse drops textDirection, textboxTightWrap, divId, cnfStyle", () => {
    // All four elements are valid pPr children that stringify emits, yet
    // parseParagraphProperties has no findChild for them — they vanish on
    // round-trip. Verified at the code level, not just the declared field set.
    const opts = parsePPr(
      `<w:textDirection w:val="lr"/>` +
        `<w:textboxTightWrap w:val="none"/>` +
        `<w:divId w:val="1"/>` +
        `<w:cnfStyle w:val="000000010000"/>`,
    );
    expect(opts.textDirection).toBeUndefined();
    expect(opts.textboxTightWrap).toBeUndefined();
    expect(opts.divId).toBeUndefined();
    expect(opts.cnfStyle).toBeUndefined();
  });

  it("F3 — parse drops revision (w:pPrChange)", () => {
    const opts = parsePPr(
      `<w:pPrChange w:id="1" w:author="a" w:date="2024-01-01T00:00:00Z">` +
        `<w:pPr><w:jc w:val="left"/></w:pPr>` +
        `</w:pPrChange>`,
    );
    expect(opts.revision).toBeUndefined();
  });

  it("declared F3 parse-loss matches the live parse gap (regression guard)", () => {
    // If parseParagraphProperties gains a findChild for any of these, the
    // field set above must be updated too — this keeps FIELD_SPECS honest.
    const report = diffTagSets(spec);
    // Order follows writeFields (textboxTightWrap precedes textDirection there).
    expect(report.f3ParseLoss).toEqual([
      "textboxTightWrap",
      "textDirection",
      "divId",
      "cnfStyle",
      "revision",
    ]);
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
