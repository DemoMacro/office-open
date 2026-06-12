import { parse as parseXml } from "@office-open/xml";
import { describe, it, expect } from "vite-plus/test";

import { stringify, parse } from "../../descriptor";
import type { BevelOptions } from "./bevel";
import type { Scene3DOptions } from "./scene-3d";
import type { Shape3DOptions } from "./shape-3d";
import { bevelDesc, shape3DDesc, scene3DDesc } from "./three-d-descriptors";

function roundTrip<T>(desc: any, opts: T): T {
  const xml = stringify(desc, opts, {} as any);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return parse(desc, el, {} as any);
}

describe("bevelDesc", () => {
  it("round-trips bevel with all attributes", () => {
    const opts: BevelOptions = {
      w: 76200,
      h: 76200,
      prst: "circle",
    };
    const result = roundTrip(bevelDesc, opts);
    expect(result.w).toBe(76200);
    expect(result.h).toBe(76200);
    expect(result.prst).toBe("circle");
  });

  it("round-trips bevel with only preset", () => {
    const opts: BevelOptions = { prst: "softRound" };
    const result = roundTrip(bevelDesc, opts);
    expect(result.prst).toBe("softRound");
  });
});

describe("shape3DDesc", () => {
  it("round-trips basic 3D shape with bevels", () => {
    const opts: Shape3DOptions = {
      bevelT: { w: 76200, h: 76200, prst: "circle" },
      bevelB: { w: 38100, h: 38100, prst: "relaxedInset" },
    };
    const result = roundTrip(shape3DDesc, opts);
    expect(result.bevelT).toBeDefined();
    expect(result.bevelT!.w).toBe(76200);
    expect(result.bevelT!.h).toBe(76200);
    expect(result.bevelT!.prst).toBe("circle");
    expect(result.bevelB).toBeDefined();
    expect(result.bevelB!.w).toBe(38100);
    expect(result.bevelB!.prst).toBe("relaxedInset");
  });

  it("round-trips 3D shape with colors and attributes", () => {
    const opts: Shape3DOptions = {
      z: 127000,
      extrusionH: 76200,
      contourW: 25400,
      prstMaterial: "matte",
      extrusionColor: { value: "4472C4" },
      contourColor: { value: "333333" },
    };
    const result = roundTrip(shape3DDesc, opts);
    expect(result.z).toBe(127000);
    expect(result.extrusionH).toBe(76200);
    expect(result.contourW).toBe(25400);
    expect(result.prstMaterial).toBe("matte");
    expect(result.extrusionColor).toEqual({ value: "4472C4" });
    expect(result.contourColor).toEqual({ value: "333333" });
  });
});

describe("scene3DDesc", () => {
  it("round-trips basic scene with camera and lightRig", () => {
    const opts: Scene3DOptions = {
      camera: { preset: "perspectiveFront" },
      lightRig: { rig: "threePt", direction: "t" },
    };
    const result = roundTrip(scene3DDesc, opts);
    expect(result.camera).toBeDefined();
    expect(result.camera.preset).toBe("perspectiveFront");
    expect(result.lightRig).toBeDefined();
    expect(result.lightRig.rig).toBe("threePt");
    expect(result.lightRig.direction).toBe("t");
  });

  it("round-trips camera with rotation", () => {
    const opts: Scene3DOptions = {
      camera: {
        preset: "isometricTopUp",
        fov: 3600000,
        zoom: "100%",
        rotation: { lat: 0, lon: 0, rev: 5400000 },
      },
      lightRig: { rig: "balanced", direction: "tl" },
    };
    const result = roundTrip(scene3DDesc, opts);
    expect(result.camera.preset).toBe("isometricTopUp");
    expect(result.camera.fov).toBe(3600000);
    expect(result.camera.zoom).toBe("100%");
    expect(result.camera.rotation).toEqual({ lat: 0, lon: 0, rev: 5400000 });
  });

  it("round-trips lightRig with rotation", () => {
    const opts: Scene3DOptions = {
      camera: { preset: "perspectiveFront" },
      lightRig: {
        rig: "soft",
        direction: "b",
        rotation: { lat: 5400000, lon: 0, rev: 0 },
      },
    };
    const result = roundTrip(scene3DDesc, opts);
    expect(result.lightRig.rig).toBe("soft");
    expect(result.lightRig.direction).toBe("b");
    expect(result.lightRig.rotation).toEqual({ lat: 5400000, lon: 0, rev: 0 });
  });

  it("round-trips scene with backdrop", () => {
    const opts: Scene3DOptions = {
      camera: { preset: "perspectiveFront" },
      lightRig: { rig: "threePt", direction: "t" },
      backdrop: {
        anchor: { x: 0, y: 0, z: 0 },
        normal: { dx: 0, dy: 0, dz: 1 },
        up: { dx: 0, dy: 1, dz: 0 },
      },
    };
    const result = roundTrip(scene3DDesc, opts);
    expect(result.backdrop).toBeDefined();
    expect(result.backdrop!.anchor).toEqual({ x: 0, y: 0, z: 0 });
    expect(result.backdrop!.normal).toEqual({ dx: 0, dy: 0, dz: 1 });
    expect(result.backdrop!.up).toEqual({ dx: 0, dy: 1, dz: 0 });
  });
});
