import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseParagraphProperties } from "../../body";
import type { DocxReadContext } from "../../context";
import { stringifyParagraphProperties } from "./stringify";

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const readCtx = {} as unknown as DocxReadContext;

function roundTrip(props: Record<string, unknown>): Record<string, unknown> {
  const { xml } = stringifyParagraphProperties(props as never);
  const doc = parseXml(xml!.replace("<w:pPr>", `<w:pPr ${W_NS}>`));
  const pPr = doc.elements?.[0];
  if (!pPr) throw new Error("parsed document has no root element");
  return parseParagraphProperties(pPr, readCtx) as Record<string, unknown>;
}

describe("paragraph properties measure round-trip", () => {
  // UniversalMeasure is input-only: stringify normalizes it to twip integers (matching
  // MS Office's integer-only output), so a UM value round-trips as its twip equivalent.
  it("normalizes spacing before/after UniversalMeasure (mm) to twips", () => {
    const result = roundTrip({ spacing: { before: "1.5mm", after: "2mm" } });
    const spacing = result.spacing as Record<string, unknown>;
    expect(spacing.before).toBe(85);
    expect(spacing.after).toBe(113);
  });

  it("normalizes spacing.line UniversalMeasure (mm) to twips", () => {
    const result = roundTrip({ spacing: { line: "3mm", lineRule: "exact" } });
    const spacing = result.spacing as Record<string, unknown>;
    expect(spacing.line).toBe(170);
  });

  it("normalizes indent left/firstLine UniversalMeasure (mm) to twips", () => {
    const result = roundTrip({ indent: { left: "5mm", firstLine: "2.5mm" } });
    const indent = result.indent as Record<string, unknown>;
    expect(indent.left).toBe(283);
    expect(indent.firstLine).toBe(141);
  });

  it("normalizes indent right/hanging UniversalMeasure (mm) to twips", () => {
    const result = roundTrip({ indent: { right: "4mm", hanging: "1mm" } });
    const indent = result.indent as Record<string, unknown>;
    expect(indent.right).toBe(226);
    expect(indent.hanging).toBe(56);
  });
});
