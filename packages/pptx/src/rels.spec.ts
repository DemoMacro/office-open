import { unzipSync } from "@office-open/core";
import { describe, expect, it } from "vite-plus/test";

import { generatePresentation } from "./generate";
import { parsePresentation } from "./parse";
import type { PresentationOptions } from "./shared/file";

// Slide rels mix model allocations (layout, media, …) with passthrough
// re-emission at source ids (verbatim slide islands reference them). The
// source id space is reserved up front, so a media batch on an edited
// round-trip slide can never take an id a source re-use needs — a collision
// would either duplicate the id (package refused by Office) or force the
// claim to renumber (the verbatim r:id reference dangles).

const CHART_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const LAYOUT_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const MEDIA_REL = "http://schemas.microsoft.com/office/2007/relationships/media";
const PACKAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
const THEME_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

const decodeEntry = (buffer: Uint8Array, path: string): string => {
  const unzipped = unzipSync(buffer);
  const entry = unzipped[path];
  if (!entry) throw new Error(`missing zip entry: ${path}`);
  return new TextDecoder().decode(entry);
};

// Source slide1.xml.rels: rId1 layout, rId2 a chart part the verbatim island
// still references by r:id. The edit adds a modeled picture — without the
// reserve the media batch takes rId2 and the chart claim renumbers.
describe("slide rels with passthrough source ids and a modeled picture", () => {
  it("keeps source ids for re-used rels and allocates the picture above them", async () => {
    const options: PresentationOptions = {
      slides: [
        {
          children: [
            {
              picture: {
                type: "png",
                x: 0,
                y: 0,
                width: 100,
                height: 100,
                data: "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
              },
            },
          ],
        },
      ],
      passthroughRelationships: [
        {
          source: "ppt/slides/slide1.xml",
          relationshipType: LAYOUT_REL,
          target: "../slideLayouts/slideLayout1.xml",
          rId: "rId1",
        },
        {
          source: "ppt/slides/slide1.xml",
          relationshipType: CHART_REL,
          target: "../charts/chart1.xml",
          rId: "rId2",
        },
      ],
      rawParts: [{ path: "ppt/charts/chart1.xml", data: "<c:chartSpace/>" }],
    };

    const buffer = await generatePresentation(options);
    const rels = decodeEntry(buffer, "ppt/slides/_rels/slide1.xml.rels");

    const ids = [...rels.matchAll(/Id="rId(\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
    // The verbatim chart reference keeps its source id
    expect(rels).toMatch(new RegExp(`Id="rId2"[^>]*Type="${CHART_REL}"`));
    // The fresh picture lands above the source id space
    const imageId = Number(/Id="rId(\d+)"[^>]*relationships\/image"/.exec(rels)?.[1]);
    expect(imageId).toBeGreaterThan(2);
  });

  it("keeps media relationships with different targets of the same kind", async () => {
    const options: PresentationOptions = {
      slides: [
        {
          children: [
            {
              picture: {
                type: "png",
                fileName: "kept.png",
                x: 0,
                y: 0,
                width: 100,
                height: 100,
                data: "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
              },
            },
          ],
        },
      ],
      passthroughRelationships: [
        {
          source: "ppt/slides/slide1.xml",
          relationshipType: IMAGE_REL,
          target: "../media/kept.png",
          rId: "rId2",
        },
        {
          source: "ppt/slides/slide1.xml",
          relationshipType: IMAGE_REL,
          target: "../media/legacy.wmf",
          rId: "rId3",
        },
      ],
      rawParts: [{ path: "ppt/media/legacy.wmf", data: "legacy" }],
    };

    const buffer = await generatePresentation(options);
    const rels = decodeEntry(buffer, "ppt/slides/_rels/slide1.xml.rels");

    expect(rels).toContain('Target="../media/kept.png"');
    expect(rels).toContain('Id="rId3"');
    expect(rels).toContain('Target="../media/legacy.wmf"');
  });

  it("keeps duplicate source image relationships to one media target", async () => {
    const options: PresentationOptions = {
      slides: [
        {
          children: [
            {
              picture: {
                type: "png",
                fileName: "same.png",
                x: 0,
                y: 0,
                width: 100,
                height: 100,
                data: "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
              },
            },
          ],
        },
      ],
      passthroughRelationships: [
        {
          source: "ppt/slides/slide1.xml",
          relationshipType: IMAGE_REL,
          target: "../media/same.png",
          rId: "rId3",
        },
        {
          source: "ppt/slides/slide1.xml",
          relationshipType: IMAGE_REL,
          target: "../media/same.png",
          rId: "rId2",
        },
      ],
    };

    const buffer = await generatePresentation(options);
    const rels = decodeEntry(buffer, "ppt/slides/_rels/slide1.xml.rels");
    const imageTargets = [
      ...rels.matchAll(/Id="(rId\d+)"[^>]*Type="([^"]+)"[^>]*Target="([^"]+)"/g),
    ].filter((match) => match[2] === IMAGE_REL && match[3] === "../media/same.png");

    expect(imageTargets.map((match) => match[1])).toEqual(["rId2", "rId3"]);
  });

  it("keeps modeled video media relationship ids stable across round-trip", async () => {
    const options: PresentationOptions = {
      slides: [
        {
          children: [
            {
              video: {
                x: 0,
                y: 0,
                width: 100,
                height: 100,
                name: "Test Video",
                data: new Uint8Array([0, 0, 0, 24, 102, 116, 121, 112]),
                type: "mp4",
                poster:
                  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
                posterType: "png",
              },
            },
          ],
        },
      ],
    };
    const buffer = await generatePresentation(options);
    const parsed = await parsePresentation(buffer);
    const roundTripped = await generatePresentation(parsed);
    const rels = decodeEntry(roundTripped, "ppt/slides/_rels/slide1.xml.rels");

    const relationships = [
      ...rels.matchAll(/Id="(rId\d+)"[^>]*Type="([^"]+)"[^>]*Target="([^"]+)"/g),
    ].map((match) => [match[1], match[2], match[3]] as const);
    expect(relationships).toEqual([
      [
        "rId1",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
        "../slideLayouts/slideLayout1.xml",
      ],
      ["rId2", IMAGE_REL, "../media/Test_Video_poster.png"],
      ["rId3", MEDIA_REL, "../media/Test_Video.mp4"],
      [
        "rId4",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video",
        "../media/Test_Video.mp4",
      ],
    ]);
  });
});

// Captured rels whose kind the model re-registers (the master theme, say)
// are absorbed — the claim skips them as owned, so their ids must not be
// reserved: a reservation opens a hole the round-trip then reports as drift
// (theme re-registered above the reserved id, the captured slot unused).
describe("captured rels absorbed by the model", () => {
  it("does not reserve ids for kinds the compiler re-emits, keeping round-trip stable", async () => {
    const options: PresentationOptions = {
      slides: [{ children: [] }],
      passthroughRelationships: [
        {
          source: "ppt/slideMasters/slideMaster1.xml",
          relationshipType: THEME_REL,
          target: "../theme/theme1.xml",
          rId: "rId2",
        },
      ],
    };

    const buffer = await generatePresentation(options);
    const rels = decodeEntry(buffer, "ppt/slideMasters/_rels/slideMaster1.xml.rels");

    const ids = [...rels.matchAll(/Id="rId(\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
    // The model's theme registration lands at the captured source id — no
    // rId2 hole pushing it to rId3
    expect(rels).toMatch(new RegExp(`Id="rId2"[^>]*Type="${THEME_REL}"`));
    expect(rels).not.toMatch(/Id="rId3"/);
  });
});

describe("OLE embeddings owned by the model", () => {
  it("reserves package-embedding source ids before media allocation", async () => {
    const options: PresentationOptions = {
      slides: [
        {
          children: [
            {
              picture: {
                type: "png",
                x: 0,
                y: 0,
                width: 100,
                height: 100,
                data: "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
              },
            },
            {
              ole: {
                x: 0,
                y: 0,
                width: 100,
                height: 100,
                progId: "Word.Document.12",
                embed: { data: "Zm9v", relationshipType: "package" },
              },
            },
          ],
        },
      ],
      passthroughRelationships: [
        {
          source: "ppt/slides/slide1.xml",
          relationshipType: PACKAGE_REL,
          target: "../embeddings/oleObject1.bin",
          rId: "rId2",
        },
      ],
    };

    const buffer = await generatePresentation(options);
    const slide = decodeEntry(buffer, "ppt/slides/slide1.xml");
    const rels = decodeEntry(buffer, "ppt/slides/_rels/slide1.xml.rels");

    expect(slide).toContain('r:id="rId2"');
    expect(rels).toMatch(new RegExp(`Id="rId2"[^>]*Type="${PACKAGE_REL}"`));
    const imageId = Number(/Id="rId(\d+)"[^>]*relationships\/image"/.exec(rels)?.[1]);
    expect(imageId).toBeGreaterThan(2);
  });
});
