import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { lockedCanvasDesc } from "./locked-canvas";
import type { LockedCanvasDescriptorOptions } from "./locked-canvas";

// ── Mock contexts ──

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: LockedCanvasDescriptorOptions) {
  const xml = lockedCanvasDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return lockedCanvasDesc.parse(el, readCtx);
}

describe("lockedCanvasDesc round-trip", () => {
  it("round-trips basic position and name", () => {
    const opts: LockedCanvasDescriptorOptions = {
      id: 100,
      name: "Test Canvas",
      x: 50,
      y: 60,
      width: 200,
      height: 150,
    };
    const result = roundTrip(opts);

    expect(result.id).toBe(100);
    expect(result.name).toBe("Test Canvas");
    expect(result.x).toBe(50);
    expect(result.y).toBe(60);
    expect(result.width).toBe(200);
    expect(result.height).toBe(150);
  });

  it("round-trips with children shapes (excluding textBody)", () => {
    const opts: LockedCanvasDescriptorOptions = {
      id: 200,
      name: "Canvas with Children",
      x: 10,
      y: 20,
      width: 400,
      height: 300,
      children: [
        {
          x: 5,
          y: 10,
          width: 100,
          height: 50,
          geometry: "ellipse",
          fill: "FF0000",
        },
        {
          x: 120,
          y: 10,
          width: 80,
          height: 40,
          geometry: "roundRect",
          fill: "00FF00",
        },
      ],
    };
    const result = roundTrip(opts);

    expect(result.id).toBe(200);
    expect(result.name).toBe("Canvas with Children");
    expect(result.x).toBe(10);
    expect(result.y).toBe(20);
    expect(result.width).toBe(400);
    expect(result.height).toBe(300);
    expect(result.children).toBeDefined();
    expect(result.children).toHaveLength(2);

    const child1 = result.children![0];
    expect(child1.x).toBe(5);
    expect(child1.y).toBe(10);
    expect(child1.width).toBe(100);
    expect(child1.height).toBe(50);
    expect(child1.geometry).toBe("ellipse");
    expect(child1.fill).toBe("FF0000");

    const child2 = result.children![1];
    expect(child2.x).toBe(120);
    expect(child2.y).toBe(10);
    expect(child2.width).toBe(80);
    expect(child2.height).toBe(40);
    expect(child2.geometry).toBe("roundRect");
    expect(child2.fill).toBe("00FF00");
  });

  it("round-trips without children", () => {
    const opts: LockedCanvasDescriptorOptions = {
      id: 300,
      name: "Empty Canvas",
      x: 0,
      y: 0,
      width: 500,
      height: 400,
    };
    const result = roundTrip(opts);

    expect(result.id).toBe(300);
    expect(result.name).toBe("Empty Canvas");
    expect(result.x).toBe(0);
    expect(result.y).toBe(0);
    expect(result.width).toBe(500);
    expect(result.height).toBe(400);
    expect(result.children).toBeUndefined();
  });

  it("round-trips default geometry for child shape", () => {
    const opts: LockedCanvasDescriptorOptions = {
      id: 400,
      children: [
        {
          x: 10,
          y: 10,
          width: 50,
          height: 50,
          // geometry defaults to "rect"
        },
      ],
    };
    const result = roundTrip(opts);

    expect(result.children).toHaveLength(1);
    expect(result.children![0].geometry).toBe("rect");
  });

  it("round-trips EMU conversion correctly", () => {
    const opts: LockedCanvasDescriptorOptions = {
      id: 500,
      x: 1024,
      y: 768,
      width: 1920,
      height: 1080,
    };
    const result = roundTrip(opts);

    expect(result.x).toBe(1024);
    expect(result.y).toBe(768);
    expect(result.width).toBe(1920);
    expect(result.height).toBe(1080);
  });
});
