import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { chartsheetDesc } from "./chartsheet";
import type { ChartsheetDescriptorOptions } from "./chartsheet";

// ── Minimal context stubs ──

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: ChartsheetDescriptorOptions) {
  const xml = chartsheetDesc.stringify(opts, writeCtx)!;
  // nativeTypeAttributes mirrors the real xlsx parse path (ParsedArchive.get
  // coerces "1"/"0" to numbers), so boolean reads are exercised against
  // numeric coercion rather than a permissive non-coerced parse.
  const doc = parseXml(xml, { nativeTypeAttributes: true });
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return chartsheetDesc.parse(el, readCtx) as unknown as ChartsheetDescriptorOptions;
}

// ── Tests ──

describe("chartsheetDesc round-trip", () => {
  const baseOpts: ChartsheetDescriptorOptions = {
    drawingRId: "rId1",
    chart: {
      type: "column",
      series: [{ name: "Sales", values: [10, 20, 30] }],
    },
  };

  it("round-trips minimal chartsheet", () => {
    const result = roundTrip(baseOpts);

    expect(result.zoomToFit).toBeUndefined();
    expect(result.published).toBeUndefined();
  });

  it("round-trips zoomToFit", () => {
    const opts: ChartsheetDescriptorOptions = { ...baseOpts, zoomToFit: true };
    const result = roundTrip(opts);

    expect(result.zoomToFit).toBe(true);
  });

  it("round-trips published flag (XSD default true — only false survives)", () => {
    const opts: ChartsheetDescriptorOptions = { ...baseOpts, published: false };
    const result = roundTrip(opts);

    expect(result.published).toBe(false);
  });

  it("round-trips tabColor", () => {
    const opts: ChartsheetDescriptorOptions = { ...baseOpts, tabColor: "FF4472C4" };
    const result = roundTrip(opts);

    expect(result.tabColor).toBe("FF4472C4");
  });

  it("round-trips pageMargins", () => {
    const opts: ChartsheetDescriptorOptions = {
      ...baseOpts,
      pageMargins: { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
    };
    const result = roundTrip(opts);

    expect(result.pageMargins?.left).toBe(0.5);
    expect(result.pageMargins?.top).toBe(0.75);
    expect(result.pageMargins?.footer).toBe(0.3);
  });

  it("round-trips pageSetup", () => {
    const opts: ChartsheetDescriptorOptions = {
      ...baseOpts,
      pageSetup: { paperSize: 9, orientation: "landscape", horizontalDpi: 300, copies: 2 },
    };
    const result = roundTrip(opts);

    expect(result.pageSetup?.paperSize).toBe(9);
    expect(result.pageSetup?.orientation).toBe("landscape");
    expect(result.pageSetup?.horizontalDpi).toBe(300);
    expect(result.pageSetup?.copies).toBe(2);
  });

  it("round-trips headerFooter", () => {
    const opts: ChartsheetDescriptorOptions = {
      ...baseOpts,
      headerFooter: {
        differentFirst: true,
        differentOddEven: true,
        oddHeader: "Page",
        oddFooter: "End",
      },
    };
    const result = roundTrip(opts);

    expect(result.headerFooter?.differentFirst).toBe(true);
    expect(result.headerFooter?.differentOddEven).toBe(true);
    expect(result.headerFooter?.oddHeader).toBe("Page");
    expect(result.headerFooter?.oddFooter).toBe("End");
  });

  it("round-trips sheetProtection flags under nativeTypeAttributes coercion", () => {
    const opts: ChartsheetDescriptorOptions = {
      ...baseOpts,
      sheetProtection: { content: true, objects: true },
    };
    const result = roundTrip(opts);

    expect(result.sheetProtection?.content).toBe(true);
    expect(result.sheetProtection?.objects).toBe(true);
  });

  it("round-trips codeName", () => {
    const opts: ChartsheetDescriptorOptions = { ...baseOpts, codeName: "Sheet1Code" };
    const result = roundTrip(opts);

    expect(result.codeName).toBe("Sheet1Code");
  });
});
