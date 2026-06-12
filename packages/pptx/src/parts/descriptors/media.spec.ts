import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { videoDesc, audioDesc } from "./media";
import type { VideoDescriptorOptions, AudioDescriptorOptions } from "./media";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTripVideo(opts: VideoDescriptorOptions) {
  const xml = videoDesc.stringify(opts, writeCtx);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return videoDesc.parse(el, readCtx);
}

function roundTripAudio(opts: AudioDescriptorOptions) {
  const xml = audioDesc.stringify(opts, writeCtx);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return audioDesc.parse(el, readCtx);
}

describe("videoDesc round-trip", () => {
  it("round-trips basic video position and name", () => {
    const opts: VideoDescriptorOptions = {
      id: 101,
      name: "My Video",
      x: 50,
      y: 100,
      width: 640,
      height: 480,
    };
    const result = roundTripVideo(opts);

    expect(result.name).toBe("My Video");
    expect(result.x).toBeCloseTo(50, 0);
    expect(result.y).toBeCloseTo(100, 0);
    expect(result.width).toBeCloseTo(640, 0);
    expect(result.height).toBeCloseTo(480, 0);
  });

  it("round-trips video with default name", () => {
    const opts: VideoDescriptorOptions = {
      id: 102,
    };
    const result = roundTripVideo(opts);

    // Default name is "Video 102"
    expect(result.name).toBe("Video 102");
  });

  it("round-trips video with zero dimensions", () => {
    const opts: VideoDescriptorOptions = {
      id: 103,
      name: "Placeholder",
      x: 0,
      y: 0,
      width: 0,
      height: 0,
    };
    const result = roundTripVideo(opts);

    expect(result.name).toBe("Placeholder");
    expect(result.x).toBeCloseTo(0, 0);
    expect(result.y).toBeCloseTo(0, 0);
  });

  // Note: data/poster are not round-tripped because the read context
  // cannot resolve media relationships in this test setup.
});

describe("audioDesc round-trip", () => {
  it("round-trips basic audio position and name", () => {
    const opts: AudioDescriptorOptions = {
      id: 201,
      name: "My Audio",
      x: 30,
      y: 40,
      width: 50,
      height: 50,
    };
    const result = roundTripAudio(opts);

    expect(result.name).toBe("My Audio");
    expect(result.x).toBeCloseTo(30, 0);
    expect(result.y).toBeCloseTo(40, 0);
    expect(result.width).toBeCloseTo(50, 0);
    expect(result.height).toBeCloseTo(50, 0);
  });

  it("round-trips audio with default name", () => {
    const opts: AudioDescriptorOptions = {
      id: 202,
    };
    const result = roundTripAudio(opts);

    // Default name is "Audio 202"
    expect(result.name).toBe("Audio 202");
  });

  // Note: data is not round-tripped because the read context
  // cannot resolve media relationships in this test setup.
});
