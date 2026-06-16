import { parse as parseXml } from "@office-open/xml";
import { describe, it, expect } from "vite-plus/test";

import { stringify, parse } from "../../descriptor";
import type { BlipOptions } from "./blip";
import {
  blipDesc,
  blipFillDesc,
  tileDesc,
  sourceRectangleDesc,
  stretchDesc,
} from "./blip-descriptors";
import type { SourceRectangleOptions } from "./source-rectangle";
import type { TileOptions } from "./tile";

function roundTrip<T>(desc: any, opts: T): T {
  const xml = stringify(desc, opts, {} as any);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return parse(desc, el, {} as any);
}

describe("tileDesc", () => {
  it("round-trips tile with all attributes", () => {
    const opts: TileOptions = {
      tx: 100,
      ty: 200,
      sx: 50000,
      sy: 50000,
      flip: "x",
      align: "center",
    };
    const result = roundTrip(tileDesc, opts);
    expect(result.tx).toBe(100);
    expect(result.ty).toBe(200);
    expect(result.sx).toBe(50000);
    expect(result.sy).toBe(50000);
    expect(result.flip).toBe("x");
    expect(result.align).toBe("center");
  });
});

describe("sourceRectangleDesc", () => {
  it("round-trips source rectangle", () => {
    const opts: SourceRectangleOptions = {
      left: 10,
      top: 20,
      right: 30,
      bottom: 40,
    };
    const result = roundTrip(sourceRectangleDesc, opts);
    expect(result.left).toBe(10);
    expect(result.top).toBe(20);
    expect(result.right).toBe(30);
    expect(result.bottom).toBe(40);
  });
});

describe("stretchDesc", () => {
  it("round-trips stretch with fill rect", () => {
    const opts: SourceRectangleOptions = {
      left: 5,
      top: 5,
      right: 5,
      bottom: 5,
    };
    const result = roundTrip(stretchDesc, opts);
    expect(result.left).toBe(5);
    expect(result.top).toBe(5);
    expect(result.right).toBe(5);
    expect(result.bottom).toBe(5);
  });
});

describe("blipDesc", () => {
  it("round-trips blip with referenceId", () => {
    type BlipFull = BlipOptions & { blipEffects?: any };
    const opts: BlipFull = { referenceId: "image1.png" };
    const result = roundTrip(blipDesc, opts);
    expect(result.referenceId).toBe("image1.png");
  });

  it("round-trips blip with grayscale effect", () => {
    type BlipFull = BlipOptions & { blipEffects?: any };
    const opts: BlipFull = {
      referenceId: "img.png",
      blipEffects: { grayscale: true },
    };
    const result = roundTrip(blipDesc, opts);
    expect(result.referenceId).toBe("img.png");
    expect(result.blipEffects).toBeDefined();
    expect(result.blipEffects!.grayscale).toBe(true);
  });

  it("round-trips blip with luminance effect", () => {
    type BlipFull = BlipOptions & { blipEffects?: any };
    const opts: BlipFull = {
      referenceId: "img.png",
      blipEffects: { luminance: { bright: 20, contrast: 10 } },
    };
    const result = roundTrip(blipDesc, opts);
    expect(result.blipEffects!.luminance).toBeDefined();
    expect(result.blipEffects!.luminance.bright).toBe(20);
    expect(result.blipEffects!.luminance.contrast).toBe(10);
  });
});

describe("blipFillDesc", () => {
  it("round-trips blip fill with referenceId", () => {
    type BlipFillFull = {
      referenceId?: string;
      blipEffects?: any;
      dpi?: number;
      rotWithShape?: boolean;
      sourceRectangle?: SourceRectangleOptions;
      tile?: TileOptions;
    };
    const opts: BlipFillFull = {
      referenceId: "image1.png",
      dpi: 150,
      rotWithShape: true,
    };
    const result = roundTrip(blipFillDesc, opts);
    expect(result.referenceId).toBe("image1.png");
    expect(result.dpi).toBe(150);
    expect(result.rotWithShape).toBe(true);
  });

  it("round-trips blip fill with source rectangle", () => {
    type BlipFillFull = {
      referenceId?: string;
      blipEffects?: any;
      sourceRectangle?: SourceRectangleOptions;
      tile?: TileOptions;
    };
    const opts: BlipFillFull = {
      referenceId: "img.png",
      sourceRectangle: { left: 10, top: 20, right: 30, bottom: 40 },
    };
    const result = roundTrip(blipFillDesc, opts);
    expect(result.sourceRectangle).toBeDefined();
    expect(result.sourceRectangle!.left).toBe(10);
    expect(result.sourceRectangle!.top).toBe(20);
  });

  it("round-trips blip fill with tile", () => {
    type BlipFillFull = { referenceId?: string; blipEffects?: any; tile?: TileOptions };
    const opts: BlipFillFull = {
      referenceId: "img.png",
      tile: { tx: 100, ty: 200, sx: 50000, sy: 50000 },
    };
    const result = roundTrip(blipFillDesc, opts);
    expect(result.tile).toBeDefined();
    expect(result.tile!.tx).toBe(100);
    expect(result.tile!.ty).toBe(200);
  });
});
