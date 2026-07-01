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

  it("deduplicates byte-identical content across distinct buffer objects", () => {
    const media = new Media<TestEntry>();
    const first = media.addMedia(
      new Uint8Array([1, 2, 3]),
      "png",
      build(new Uint8Array([1, 2, 3]), "png"),
    );
    // A separately-allocated buffer carrying the same bytes (e.g. a re-decoded round-trip input).
    const second = media.addMedia(
      new Uint8Array([1, 2, 3]),
      "png",
      build(new Uint8Array([1, 2, 3]), "png"),
    );

    expect(second).toBe(first);
    expect(media.array).toHaveLength(1);
  });

  it("keeps byte-identical content separate when the type differs", () => {
    // mc:AlternateContent carries the same image as EMF (Choice) and WMF
    // (Fallback) — distinct content-types MS Office stores as separate parts.
    // Byte-identical dedup must not collapse them.
    const media = new Media<TestEntry>();
    const bytes = new Uint8Array([1, 2, 3]);
    const emf = media.addMedia(bytes, "emf", build(bytes, "emf"));
    const wmf = media.addMedia(bytes, "wmf", build(bytes, "wmf"));

    expect(emf.fileName).toBe("image1.emf");
    expect(wmf.fileName).toBe("image2.wmf");
    expect(emf.type).toBe("emf");
    expect(wmf.type).toBe("wmf");
    expect(media.array).toHaveLength(2);
  });

  it("does not overwrite a pinned name held by different bytes", () => {
    // Two round-trip paths may resolve to the same file name with different
    // bytes; overwriting would silently drop media. Re-allocate instead.
    const media = new Media<TestEntry>();
    const first = media.addMedia(
      new Uint8Array([1, 2, 3]),
      "png",
      build(new Uint8Array([1, 2, 3]), "png"),
      "image3.png",
    );
    const second = media.addMedia(
      new Uint8Array([9, 9, 9]),
      "png",
      build(new Uint8Array([9, 9, 9]), "png"),
      "image3.png",
    );

    expect(first.fileName).toBe("image3.png");
    expect(second.fileName).not.toBe("image3.png");
    expect(media.array).toHaveLength(2);
    // The original bytes survive under the pinned name.
    const survivor = media.array.find((e) => e.fileName === "image3.png");
    expect(survivor?.data).toEqual(new Uint8Array([1, 2, 3]));
  });

  it("skips counter-allocated names already taken", () => {
    // A pinned name occupying image1.png — the counter must skip it for new
    // bytes (mirrors round-trip where source basenames and the counter share
    // the imageN.ext namespace).
    const media = new Media<TestEntry>();
    media.addMedia(new Uint8Array([1]), "png", build(new Uint8Array([1]), "png"), "image1.png");
    const auto = media.addMedia(new Uint8Array([2]), "png", build(new Uint8Array([2]), "png"));

    expect(auto.fileName).toBe("image2.png");
    expect(media.array).toHaveLength(2);
  });
});
