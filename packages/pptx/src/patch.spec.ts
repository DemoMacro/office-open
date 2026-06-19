import { unzipSync } from "@office-open/core";
import { describe, expect, it } from "vite-plus/test";

import { generatePresentation } from "./generate";
import { parsePresentation } from "./parse";
import { patchPresentation } from "./patch";
import type { PresentationOptions, SlideOptions } from "./shared/file";

const decodeEntry = (buffer: Uint8Array, path: string): string => {
  const unzipped = unzipSync(buffer);
  const entry = unzipped[path];
  if (!entry) throw new Error(`missing zip entry: ${path}`);
  return new TextDecoder().decode(entry);
};

const slide = (text: string): SlideOptions => ({
  children: [{ shape: { x: 0, y: 0, width: 200, height: 100, textBody: { text } } }],
});

const template = (slides: SlideOptions[]): PresentationOptions => ({ slides });

describe("patchPresentation slides", () => {
  it("appends a slide and wires sldIdLst, rels, and content types", async () => {
    const buffer = await generatePresentation(template([slide("A"), slide("B")]));
    const patched = await patchPresentation({
      data: buffer,
      outputType: "uint8array",
      slides: { append: [slide("APPENDED_SLIDE")] },
    });

    // Parser must accept the new file and see three slides.
    const parsed = parsePresentation(patched);
    expect(parsed.slides?.length).toBe(3);

    // New slide part carries the appended text.
    expect(decodeEntry(patched, "ppt/slides/slide3.xml")).toContain("APPENDED_SLIDE");

    // sldIdLst grew to three entries.
    const pres = decodeEntry(patched, "ppt/presentation.xml");
    expect((pres.match(/<p:sldId /g) ?? []).length).toBe(3);

    // presentation.xml.rels references the new slide; content types declares it.
    expect(decodeEntry(patched, "ppt/_rels/presentation.xml.rels")).toContain("slides/slide3.xml");
    expect(decodeEntry(patched, "[Content_Types].xml")).toContain("/ppt/slides/slide3.xml");

    // Slide rels points at an existing layout.
    expect(decodeEntry(patched, "ppt/slides/_rels/slide3.xml.rels")).toContain("slideLayout1.xml");
  });

  it("replaces a slide by index without changing the slide count", async () => {
    const buffer = await generatePresentation(template([slide("ORIGINAL"), slide("B")]));
    const patched = await patchPresentation({
      data: buffer,
      outputType: "uint8array",
      slides: { replace: { 0: slide("REPLACED_SLIDE") } },
    });

    const parsed = parsePresentation(patched);
    expect(parsed.slides?.length).toBe(2);

    const slide1 = decodeEntry(patched, "ppt/slides/slide1.xml");
    expect(slide1).toContain("REPLACED_SLIDE");
    expect(slide1).not.toContain("<a:t>ORIGINAL</a:t>");
  });

  it("throws when replacing a non-existent index", async () => {
    const buffer = await generatePresentation(template([slide("A")]));
    await expect(
      patchPresentation({
        data: buffer,
        outputType: "uint8array",
        slides: { replace: { 5: slide("X") } },
      }),
    ).rejects.toThrow(/index 5/);
  });
});
