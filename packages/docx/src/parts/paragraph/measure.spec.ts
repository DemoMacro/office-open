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
  it("round-trips spacing before/after UniversalMeasure verbatim", () => {
    const result = roundTrip({ spacing: { before: "1.5mm", after: "2mm" } });
    const spacing = result.spacing as Record<string, unknown>;
    expect(spacing.before).toBe("1.5mm");
    expect(spacing.after).toBe("2mm");
  });

  it("round-trips spacing.line UniversalMeasure verbatim", () => {
    const result = roundTrip({ spacing: { line: "3mm", lineRule: "exact" } });
    const spacing = result.spacing as Record<string, unknown>;
    expect(spacing.line).toBe("3mm");
  });

  it("round-trips negative spacing.line signed twips", () => {
    const result = roundTrip({ spacing: { line: -240 } });
    const spacing = result.spacing as Record<string, unknown>;
    expect(spacing.line).toBe(-240);
  });

  it("round-trips indent left/firstLine UniversalMeasure verbatim", () => {
    const result = roundTrip({ indent: { left: "5mm", firstLine: "2.5mm" } });
    const indent = result.indent as Record<string, unknown>;
    expect(indent.left).toBe("5mm");
    expect(indent.firstLine).toBe("2.5mm");
  });

  it("round-trips indent right/hanging UniversalMeasure verbatim", () => {
    const result = roundTrip({ indent: { right: "4mm", hanging: "1mm" } });
    const indent = result.indent as Record<string, unknown>;
    expect(indent.right).toBe("4mm");
    expect(indent.hanging).toBe("1mm");
  });

  // hanging/firstLine are ST_TwipsMeasure (non-negative). Negative inputs
  // flip to the twin — the normalization Word itself applies when reading and
  // re-saving such (schema-invalid) markup — and lose to a legal twin value.
  it("flips negative firstLine to hanging", () => {
    const result = roundTrip({ indent: { left: 720, firstLine: -720 } });
    const indent = result.indent as Record<string, unknown>;
    expect(indent.hanging).toBe(720);
    expect(indent.firstLine).toBeUndefined();
  });

  it("flips negative hanging to firstLine", () => {
    const result = roundTrip({ indent: { hanging: -720 } });
    const indent = result.indent as Record<string, unknown>;
    expect(indent.firstLine).toBe(720);
    expect(indent.hanging).toBeUndefined();
  });

  it("keeps a legal twin over the flipped negative", () => {
    const result = roundTrip({ indent: { firstLine: -720, hanging: 567 } });
    const indent = result.indent as Record<string, unknown>;
    expect(indent.hanging).toBe(567);
    expect(indent.firstLine).toBeUndefined();
  });

  it("parses a negative w:firstLine as hanging", () => {
    const doc = parseXml(`<w:pPr ${W_NS}><w:ind w:left="720" w:firstLine="-720"/></w:pPr>`);
    const result = parseParagraphProperties(doc.elements![0]!, readCtx) as {
      indent?: Record<string, unknown>;
    };
    expect(result.indent?.hanging).toBe(720);
    expect(result.indent?.firstLine).toBeUndefined();
  });

  it("parses a negative w:hanging as firstLine", () => {
    const doc = parseXml(`<w:pPr ${W_NS}><w:ind w:hanging="-720"/></w:pPr>`);
    const result = parseParagraphProperties(doc.elements![0]!, readCtx) as {
      indent?: Record<string, unknown>;
    };
    expect(result.indent?.firstLine).toBe(720);
    expect(result.indent?.hanging).toBeUndefined();
  });
});
