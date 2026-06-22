import { parse as parseXml } from "@office-open/xml";
import { describe, it, expect } from "vite-plus/test";

import { stringify, parse } from "../descriptor";
import type { Transform2DOptions, GroupTransform2DOptions } from "./transform";
import { transform2DDesc, groupTransform2DDesc } from "./transform-descriptors";

function roundTrip<T>(desc: any, opts: T): T {
  const xml = stringify(desc, opts, {} as any);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return parse(desc, el, {} as any);
}

describe("transform2DDesc", () => {
  it("round-trips basic position and size", () => {
    const opts: Transform2DOptions = {
      x: 100000,
      y: 200000,
      width: 300000,
      height: 400000,
    };
    const result = roundTrip(transform2DDesc, opts);
    expect(result.x).toBe(100000);
    expect(result.y).toBe(200000);
    expect(result.width).toBe(300000);
    expect(result.height).toBe(400000);
  });

  it("round-trips rotation and flip", () => {
    const opts: Transform2DOptions = {
      x: 0,
      y: 0,
      width: 100000,
      height: 100000,
      rotation: 5400000,
      flipHorizontal: true,
      flipVertical: false,
    };
    const result = roundTrip(transform2DDesc, opts);
    expect(result.rotation).toBe(5400000);
    expect(result.flipHorizontal).toBe(true);
    expect(result.flipVertical).toBe(false);
  });

  it("round-trips attributes only (no children)", () => {
    const opts: Transform2DOptions = {
      rotation: 2700000,
      flipHorizontal: true,
      flipVertical: true,
    };
    const result = roundTrip(transform2DDesc, opts);
    expect(result.rotation).toBe(2700000);
    expect(result.flipHorizontal).toBe(true);
    expect(result.flipVertical).toBe(true);
  });

  it("parses a:off x/y verbatim when source uses UniversalMeasure", () => {
    // a:off x/y are ST_Coordinate (number EMU | UniversalMeasure); a source
    // emitting a measure string must be preserved, not coerced to NaN.
    const xml =
      '<a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<a:off x="5mm" y="10mm"/></a:xfrm>';
    const el = parseXml(xml).elements![0];
    const result = parse(transform2DDesc, el, {} as never);
    expect(result.x).toBe("5mm");
    expect(result.y).toBe("10mm");
  });
});

describe("groupTransform2DDesc", () => {
  it("round-trips group transform with child offsets", () => {
    const opts: GroupTransform2DOptions = {
      x: 100000,
      y: 200000,
      width: 500000,
      height: 600000,
      childOffsetX: 10000,
      childOffsetY: 20000,
      childExtentWidth: 50000,
      childExtentHeight: 60000,
    };
    const result = roundTrip(groupTransform2DDesc, opts);
    expect(result.x).toBe(100000);
    expect(result.y).toBe(200000);
    expect(result.width).toBe(500000);
    expect(result.height).toBe(600000);
    expect(result.childOffsetX).toBe(10000);
    expect(result.childOffsetY).toBe(20000);
    expect(result.childExtentWidth).toBe(50000);
    expect(result.childExtentHeight).toBe(60000);
  });
});
