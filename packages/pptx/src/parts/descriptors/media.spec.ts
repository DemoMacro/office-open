import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import type { AudioFrameOptions } from "@shared/media/audio-frame";
import type { VideoFrameOptions } from "@shared/media/video-frame";
import { describe, expect, it } from "vite-plus/test";

import { videoDesc, audioDesc } from "./media";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
  addImage: (_key: string, entry: { fileName: string }) => entry,
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function parseRoot(xml: string | undefined) {
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return el;
}

function roundTripVideo(opts: VideoFrameOptions) {
  return videoDesc.parse(parseRoot(videoDesc.stringify(opts, writeCtx)), readCtx);
}

function roundTripAudio(opts: AudioFrameOptions) {
  return audioDesc.parse(parseRoot(audioDesc.stringify(opts, writeCtx)), readCtx);
}

describe("videoDesc round-trip", () => {
  it("round-trips basic video position and name", () => {
    const opts: VideoFrameOptions = {
      id: 101,
      data: "dummy",
      type: "mp4",
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
    expect(result.type).toBe("mp4");
  });

  it("round-trips video with default name", () => {
    const opts: VideoFrameOptions = {
      id: 102,
      data: "dummy",
      type: "mp4",
    };
    const result = roundTripVideo(opts);

    // Default name is "Video 102"
    expect(result.name).toBe("Video 102");
  });

  it("round-trips video contentType", () => {
    const result = roundTripVideo({
      id: 103,
      data: "dummy",
      type: "mp4",
      name: "Placeholder",
      contentType: "video/mp4",
      x: 0,
      y: 0,
      width: 0,
      height: 0,
    });

    expect(result.contentType).toBe("video/mp4");
  });

  it("parses a:quickTimeFile as mov video", () => {
    const el = parseRoot(
      `<p:pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
        `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
        `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
        `<p:nvPicPr><p:cNvPr id="5" name="QT"/><p:cNvPicPr/><p:nvPr>` +
        `<a:quickTimeFile r:link="rId9"/></p:nvPr></p:nvPicPr></p:pic>`,
    );
    const result = videoDesc.parse(el, readCtx);
    expect(result.type).toBe("mov");
  });

  // Note: data/poster are not round-tripped because the read context
  // cannot resolve media relationships in this test setup.
});

describe("audioDesc round-trip", () => {
  it("round-trips basic audio position and name", () => {
    const opts: AudioFrameOptions = {
      id: 201,
      data: "dummy",
      type: "mp3",
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
    expect(result.type).toBe("mp3");
  });

  it("round-trips audio with default name", () => {
    const opts: AudioFrameOptions = {
      id: 202,
      data: "dummy",
      type: "mp3",
    };
    const result = roundTripAudio(opts);

    // Default name is "Audio 202"
    expect(result.name).toBe("Audio 202");
  });

  it("round-trips audio contentType", () => {
    const result = roundTripAudio({
      id: 203,
      data: "dummy",
      type: "mp3",
      name: "Named",
      contentType: "audio/mpeg",
    });

    expect(result.contentType).toBe("audio/mpeg");
  });

  it("round-trips embedded WAV with the original file name", () => {
    const xml = audioDesc.stringify(
      { id: 204, data: "dummy", type: "wav", name: "Beep", audioFileName: "beep.wav" },
      writeCtx,
    );
    expect(xml).toContain('a:wavAudioFile r:embed="{audio:Beep.wav}" name="beep.wav"');

    const result = audioDesc.parse(parseRoot(xml), readCtx);
    expect(result.type).toBe("wav");
    expect(result.audioFileName).toBe("beep.wav");
  });

  it("round-trips CD audio track points", () => {
    const result = roundTripAudio({
      id: 205,
      name: "CD Clip",
      cd: { start: { track: 1, time: 30 }, end: { track: 3, time: 245 } },
    });

    expect(result.cd).toEqual({
      start: { track: 1, time: 30 },
      end: { track: 3, time: 245 },
    });
  });

  // Note: data is not round-tripped because the read context
  // cannot resolve media relationships in this test setup.
});
