import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { stringifyRunProperties } from "../stringify";
import type { RunPropertiesOptions } from "./properties";
import { parseRunProperties } from "./run-parse";

function roundTrip(opts: RunPropertiesOptions): RunPropertiesOptions {
  const rPr = stringifyRunProperties(opts)!;
  const doc = parseXml(`<w:r>${rPr}</w:r>`);
  const r = doc.elements?.[0];
  if (!r) throw new Error("parsed document has no root element");
  const rPrEl = r.elements?.[0];
  if (!rPrEl) throw new Error("run has no rPr element");
  return parseRunProperties(rPrEl);
}

describe("w14 text effects round-trip", () => {
  it("round-trips glow with srgb color and transform", () => {
    const result = roundTrip({
      glow: { radius: 63500, color: { srgb: { val: "4472C4", alpha: 60 } } },
    });
    expect(result.glow).toEqual({
      radius: 63500,
      color: { srgb: { val: "4472C4", alpha: 60 } },
    });
  });

  it("round-trips shadow with scheme color and geometry", () => {
    const result = roundTrip({
      shadowEffect: {
        blurRadius: 50800,
        distance: 38100,
        direction: 2700000 / 60000,
        alignment: "bottomRight",
        color: { scheme: { val: "accent1", sat: 75 } },
      },
    });
    expect(result.shadowEffect).toEqual({
      blurRadius: 50800,
      distance: 38100,
      direction: 45,
      alignment: "bottomRight",
      color: { scheme: { val: "accent1", sat: 75 } },
    });
  });

  it("round-trips reflection percentages", () => {
    const result = roundTrip({
      reflection: {
        blurRadius: 6350,
        startAlpha: 15,
        startPosition: 25,
        endAlpha: 2,
        endPosition: 90,
        fadeDirection: 90,
      },
    });
    expect(result.reflection).toEqual({
      blurRadius: 6350,
      startAlpha: 15,
      startPosition: 25,
      endAlpha: 2,
      endPosition: 90,
      fadeDirection: 90,
    });
  });

  it("round-trips textOutline with fill, dash, and join", () => {
    const result = roundTrip({
      textOutline: {
        fill: { solid: { srgb: { val: "FF0000" } } },
        dash: "dash",
        join: { miterLimit: 150 },
        width: 19050,
        cap: "round",
        compound: "single",
        alignment: "center",
      },
    });
    expect(result.textOutline).toEqual({
      fill: { solid: { srgb: { val: "FF0000" } } },
      dash: "dash",
      join: { miterLimit: 150 },
      width: 19050,
      cap: "round",
      compound: "single",
      alignment: "center",
    });
  });

  it("round-trips textFill gradient with linear shade", () => {
    const result = roundTrip({
      textFill: {
        gradient: {
          stops: [
            { position: 0, color: { srgb: { val: "FFFFFF" } } },
            { position: 100, color: { scheme: { val: "accent1" } } },
          ],
          linear: { angle: 90, scaled: true },
        },
      },
    });
    expect(result.textFill).toEqual({
      gradient: {
        stops: [
          { position: 0, color: { srgb: { val: "FFFFFF" } } },
          { position: 100, color: { scheme: { val: "accent1" } } },
        ],
        linear: { angle: 90, scaled: true },
      },
    });
  });

  it("round-trips scene3d with camera and light rig rotation", () => {
    const result = roundTrip({
      scene3d: {
        camera: { preset: "isometricOffAxis1Left" },
        lightRig: {
          rig: "threePt",
          direction: "topRight",
          rotation: { lat: 0, lon: 0, rev: 1200000 / 60000 },
        },
      },
    });
    expect(result.scene3d).toEqual({
      camera: { preset: "isometricOffAxis1Left" },
      lightRig: {
        rig: "threePt",
        direction: "topRight",
        rotation: { lat: 0, lon: 0, rev: 20 },
      },
    });
  });

  it("round-trips props3d with bevels, colors, and material", () => {
    const result = roundTrip({
      props3d: {
        topBevel: { w: 63500, h: 25400, prst: "circle" },
        extrusionColor: { srgb: { val: "808080" } },
        extrusionHeight: 571500,
        contourWidth: 34925,
        material: "matte",
      },
    });
    expect(result.props3d).toEqual({
      topBevel: { w: 63500, h: 25400, prst: "circle" },
      extrusionColor: { srgb: { val: "808080" } },
      extrusionHeight: 571500,
      contourWidth: 34925,
      material: "matte",
    });
  });

  it("round-trips OpenType features", () => {
    const result = roundTrip({
      ligatures: "standardContextual",
      numForm: "lining",
      numSpacing: "proportional",
      stylisticSets: { sets: [{ id: 1, enabled: true }, { id: 3 }] },
      cntxtAlts: true,
    });
    expect(result.ligatures).toBe("standardContextual");
    expect(result.numForm).toBe("lining");
    expect(result.numSpacing).toBe("proportional");
    expect(result.stylisticSets).toEqual({ sets: [{ id: 1, enabled: true }, { id: 3 }] });
    expect(result.cntxtAlts).toBe(true);
  });
});
