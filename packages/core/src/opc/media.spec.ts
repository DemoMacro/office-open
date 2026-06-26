import { describe, expect, it } from "vite-plus/test";

import { Media } from "./media";

interface TestEntry {
  fileName: string;
  data: Uint8Array;
  type: string;
}

const build =
  (data: Uint8Array, type: string) =>
  (fileName: string): TestEntry => ({ fileName, data, type });

describe("Media", () => {
  it("allocates a sequential file name for new content", () => {
    const media = new Media<TestEntry>();
    const entry = media.addMedia(
      new Uint8Array([1, 2, 3]),
      "png",
      build(new Uint8Array([1, 2, 3]), "png"),
    );

    expect(entry.fileName).toBe("image1.png");
    expect(media.array).toHaveLength(1);
  });

  it("deduplicates byte-identical content to the shared entry", () => {
    const media = new Media<TestEntry>();
    const bytes = new Uint8Array([1, 2, 3]);

    const first = media.addMedia(bytes, "png", build(bytes, "png"));
    const second = media.addMedia(bytes, "png", build(bytes, "png"));

    expect(second).toBe(first);
    expect(media.array).toHaveLength(1);
  });

  it("keeps distinct byte content as separate entries", () => {
    const media = new Media<TestEntry>();
    media.addMedia(new Uint8Array([1]), "png", build(new Uint8Array([1]), "png"));
    media.addMedia(new Uint8Array([2]), "png", build(new Uint8Array([2]), "png"));

    expect(media.array).toHaveLength(2);
  });

  it("pins a caller-provided file name when given", () => {
    const media = new Media<TestEntry>();
    const entry = media.addMedia(
      new Uint8Array([9]),
      "jpg",
      build(new Uint8Array([9]), "jpg"),
      "custom.jpg",
    );

    expect(entry.fileName).toBe("custom.jpg");
  });

  it("reuses the existing entry when content matches, ignoring a pinned name", () => {
    const media = new Media<TestEntry>();
    const bytes = new Uint8Array([1, 2, 3]);

    media.addMedia(bytes, "png", build(bytes, "png"));
    const reused = media.addMedia(bytes, "png", build(bytes, "png"), "other.png");

    expect(reused.fileName).toBe("image1.png");
    expect(media.array).toHaveLength(1);
  });
});
