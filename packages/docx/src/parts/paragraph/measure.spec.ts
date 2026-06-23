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
  const pPr = doc.elements![0];
  return parseParagraphProperties(pPr, readCtx) as Record<string, unknown>;
}

describe("paragraph properties measure round-trip", () => {
  it("round-trips spacing before/after with UniversalMeasure (mm)", () => {
    const result = roundTrip({ spacing: { before: "1.5mm", after: "2mm" } });
    const spacing = result.spacing as Record<string, unknown>;
    expect(spacing.before).toBe("1.5mm");
    expect(spacing.after).toBe("2mm");
  });

  it("round-trips spacing.line with UniversalMeasure (mm)", () => {
    const result = roundTrip({ spacing: { line: "3mm", lineRule: "exact" } });
    const spacing = result.spacing as Record<string, unknown>;
    expect(spacing.line).toBe("3mm");
  });

  it("round-trips indent left/firstLine with UniversalMeasure (mm)", () => {
    const result = roundTrip({ indent: { left: "5mm", firstLine: "2.5mm" } });
    const indent = result.indent as Record<string, unknown>;
    expect(indent.left).toBe("5mm");
    expect(indent.firstLine).toBe("2.5mm");
  });

  it("round-trips indent right/hanging with UniversalMeasure (mm)", () => {
    const result = roundTrip({ indent: { right: "4mm", hanging: "1mm" } });
    const indent = result.indent as Record<string, unknown>;
    expect(indent.right).toBe("4mm");
    expect(indent.hanging).toBe("1mm");
  });
});
