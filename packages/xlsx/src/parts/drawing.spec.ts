import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { drawingDesc } from "./drawing";
import type { DrawingOptions } from "./drawing";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: DrawingOptions) {
  const xml = drawingDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return drawingDesc.parse(el, readCtx);
}

describe("drawingDesc round-trip", () => {
  it("returns undefined for empty images and charts", () => {
    const xml = drawingDesc.stringify({ images: [], charts: [] }, writeCtx);
    expect(xml).toBeUndefined();
  });

  it("round-trips single image", () => {
    const opts: DrawingOptions = {
      images: [{ col: 2, row: 3, rId: "rId1" }],
    };
    const result = roundTrip(opts);
    const images = result.images!;

    expect(images).toHaveLength(1);
    expect(images[0].col).toBe(2);
    expect(images[0].row).toBe(3);
    expect(images[0].rId).toBe("rId1");
  });

  it("round-trips image with offsets", () => {
    const opts: DrawingOptions = {
      images: [{ col: 1, row: 1, colOffset: 50000, rowOffset: 25000, rId: "rId1" }],
    };
    const result = roundTrip(opts);
    const images = result.images!;

    expect(images[0].colOffset).toBe(50000);
    expect(images[0].rowOffset).toBe(25000);
  });

  it("round-trips image locksWithSheet and printsWithSheet", () => {
    const opts: DrawingOptions = {
      images: [{ col: 1, row: 1, rId: "rId1", locksWithSheet: false, printsWithSheet: false }],
    };
    const result = roundTrip(opts);
    const images = result.images!;

    expect(images[0].locksWithSheet).toBe(false);
    expect(images[0].printsWithSheet).toBe(false);
  });

  it("round-trips multiple images", () => {
    const opts: DrawingOptions = {
      images: [
        { col: 1, row: 1, rId: "rId1" },
        { col: 5, row: 10, rId: "rId2" },
      ],
    };
    const result = roundTrip(opts);
    const images = result.images!;

    expect(images).toHaveLength(2);
    expect(images[1].col).toBe(5);
    expect(images[1].row).toBe(10);
  });

  it("round-trips chart", () => {
    const opts: DrawingOptions = {
      charts: [{ col: 1, row: 1, rId: "rId3" }],
    };
    const result = roundTrip(opts);
    const charts = result.charts!;

    expect(charts).toHaveLength(1);
    expect(charts[0].rId).toBe("rId3");
  });

  it("round-trips chart with offsets", () => {
    const opts: DrawingOptions = {
      charts: [{ col: 3, row: 5, colOffset: 10000, rowOffset: 20000, rId: "rId4" }],
    };
    const result = roundTrip(opts);
    const charts = result.charts!;

    expect(charts[0].col).toBe(3);
    expect(charts[0].row).toBe(5);
    expect(charts[0].colOffset).toBe(10000);
    expect(charts[0].rowOffset).toBe(20000);
  });

  it("round-trips mixed images and charts", () => {
    const opts: DrawingOptions = {
      images: [{ col: 1, row: 1, rId: "rId1" }],
      charts: [{ col: 5, row: 5, rId: "rId2" }],
    };
    const result = roundTrip(opts);

    expect(result.images).toHaveLength(1);
    expect(result.charts).toHaveLength(1);
  });
});
