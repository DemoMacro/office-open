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
  const doc = parseXml(xml);
  const el = doc.elements![0];
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

  it("round-trips published flag", () => {
    const opts: ChartsheetDescriptorOptions = { ...baseOpts, published: true };
    const result = roundTrip(opts);

    expect(result.published).toBe(true);
  });

  it("round-trips tabColor", () => {
    const opts: ChartsheetDescriptorOptions = { ...baseOpts, tabColor: "FF4472C4" };
    const result = roundTrip(opts);

    expect(result.tabColor).toBe("FF4472C4");
  });
});
