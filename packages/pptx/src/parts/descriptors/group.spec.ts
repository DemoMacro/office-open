import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { groupShapeDesc } from "./group";
import type { GroupShapeDescriptorOptions } from "./group";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: GroupShapeDescriptorOptions) {
  const xml = groupShapeDesc.stringify(opts, writeCtx);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return groupShapeDesc.parse(el, readCtx);
}

describe("groupShapeDesc round-trip", () => {
  it("round-trips an empty group", () => {
    const opts: GroupShapeDescriptorOptions = {};
    const result = roundTrip(opts);

    // Group always has defaults for position/size
    expect(result).toBeDefined();
  });

  it("round-trips group position and size", () => {
    const opts: GroupShapeDescriptorOptions = {
      x: 100,
      y: 200,
      width: 400,
      height: 300,
    };
    const result = roundTrip(opts);

    expect(result.x).toBeCloseTo(100, 0);
    expect(result.y).toBeCloseTo(200, 0);
    expect(result.width).toBeCloseTo(400, 0);
    expect(result.height).toBeCloseTo(300, 0);
  });

  it("round-trips group with flipHorizontal", () => {
    const opts: GroupShapeDescriptorOptions = {
      x: 50,
      y: 50,
      width: 200,
      height: 200,
      flipHorizontal: true,
    };
    const result = roundTrip(opts);

    expect(result.flipHorizontal).toBe(true);
  });

  it("round-trips group with rotation", () => {
    const opts: GroupShapeDescriptorOptions = {
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      rotation: 5400000,
    };
    const result = roundTrip(opts);

    expect(result.rotation).toBe(5400000);
  });

  it("round-trips group with nested shape child", () => {
    const opts: GroupShapeDescriptorOptions = {
      x: 10,
      y: 20,
      width: 500,
      height: 400,
      children: [
        {
          shape: {
            id: 5,
            name: "Inner Shape",
            x: 30,
            y: 40,
            width: 100,
            height: 80,
          },
        },
      ],
    };
    const result = roundTrip(opts);

    expect(result.x).toBeCloseTo(10, 0);
    expect(result.y).toBeCloseTo(20, 0);
    expect(result.width).toBeCloseTo(500, 0);
    expect(result.height).toBeCloseTo(400, 0);
    expect(result.children).toBeDefined();
    expect(result.children!.length).toBeGreaterThan(0);
  });
});
