import { PatchError } from "@office-open/core";
import { unzipSync, zipSync } from "@office-open/core/opc";
import { describe, expect, it } from "vite-plus/test";

import { generatePresentation } from "./generate";
import { parsePresentationSync } from "./parse";
import { patchPresentation } from "./patch";
import type { PresentationOptions, SlideCommentOptions, SlideOptions } from "./shared/file";

const slide = (text: string, name?: string): SlideOptions => ({
  ...(name ? { name } : {}),
  children: [{ shape: { x: 0, y: 0, width: 200, height: 100, textBody: { text } } }],
});

const presentation = (slides: SlideOptions[]): PresentationOptions => ({ slides });

describe("patchPresentation commands", () => {
  it("replaces text in one addressed slide", async () => {
    const source = await generatePresentation(presentation([slide("Hello {{name}}")]));
    const patched = await patchPresentation({
      data: source,
      commands: [
        {
          type: "replace",
          at: { field: "slides", route: [{ index: 0 }], text: { placeholder: "name" } },
          value: "Alice",
        },
      ],
    });
    const parsed = parsePresentationSync(patched);
    expect(JSON.stringify(parsed.slides)).toContain("Alice");
    expect(JSON.stringify(parsed.slides)).not.toContain("{{name}}");
  });

  it("replaces every occurrence only when all is selected", async () => {
    const repeated: SlideOptions = {
      children: [
        { shape: { x: 0, y: 0, width: 200, height: 100, textBody: { text: "{{name}}" } } },
        { shape: { x: 0, y: 120, width: 200, height: 100, textBody: { text: "{{name}}" } } },
      ],
    };
    const source = await generatePresentation(presentation([repeated]));
    const parsed = parsePresentationSync(
      await patchPresentation({
        data: source,
        commands: [
          {
            type: "replace",
            at: {
              field: "slides",
              route: [{ index: 0 }],
              text: { placeholder: "name", all: true },
            },
            value: "Alice",
          },
        ],
      }),
    );
    expect(JSON.stringify(parsed.slides).match(/Alice/g)).toHaveLength(2);
  });

  it("inserts, replaces, moves, and removes slides in command order", async () => {
    const source = await generatePresentation(presentation([slide("A"), slide("B", "Summary")]));
    const parsed = parsePresentationSync(
      await patchPresentation({
        data: source,
        commands: [
          { type: "insert", at: { field: "slides" }, position: "end", value: [slide("C")] },
          {
            type: "replace",
            at: { field: "slides", route: [{ name: "Summary" }] },
            value: slide("Replaced", "Summary"),
          },
          { type: "move", at: { field: "slides", route: [{ index: 2 }] }, to: "start" },
          { type: "remove", at: { field: "slides", route: [{ index: 1 }] } },
        ],
      }),
    );

    expect(parsed.slides).toHaveLength(2);
    expect(JSON.stringify(parsed.slides)).toContain("C");
    expect(JSON.stringify(parsed.slides)).toContain("Replaced");
  });

  it("rejects an inserted slide with an existing slide ID", async () => {
    const source = await generatePresentation(presentation([slide("A")]));
    const original = parsePresentationSync(source).slides![0]!;

    await expect(
      patchPresentation({
        data: source,
        commands: [
          {
            type: "insert",
            at: { field: "slides" },
            position: "end",
            value: [{ ...slide("B"), slideId: original.slideId }],
          },
        ],
      }),
    ).rejects.toThrow(`Duplicate presentation slide ID: ${original.slideId}`);
  });

  it("preserves custom-show slides when patching a relocated presentation part", async () => {
    const slides = [slide("A"), slide("B")];
    const initial = await generatePresentation(presentation(slides));
    const relationships = new TextDecoder().decode(
      unzipSync(initial)["ppt/_rels/presentation.xml.rels"]!,
    );
    const slideRelationshipIds = [
      ...relationships.matchAll(
        /<Relationship\b(?=[^>]*\bId="([^"]+)")(?=[^>]*\bType="[^"]*\/slide")(?=[^>]*\bTarget="slides\/slide\d+\.xml")[^>]*\/>/g,
      ),
    ].map((match) => match[1]!);
    expect(slideRelationshipIds).toHaveLength(2);

    const source = await generatePresentation({
      ...presentation(slides),
      customShows: [
        {
          name: "Keep first",
          id: 1,
          slides: slideRelationshipIds.map((rId) => ({ rId })),
        },
      ],
    });
    const archive = unzipSync(source);
    archive["custom/deck.xml"] = archive["ppt/presentation.xml"]!;
    archive["custom/_rels/deck.xml.rels"] = new TextEncoder().encode(
      new TextDecoder()
        .decode(archive["ppt/_rels/presentation.xml.rels"]!)
        .replace(/Target="([^"]+)"/g, (_match, target: string) => `Target="../ppt/${target}"`),
    );
    archive["_rels/.rels"] = new TextEncoder().encode(
      new TextDecoder()
        .decode(archive["_rels/.rels"]!)
        .replace('Target="ppt/presentation.xml"', 'Target="custom/deck.xml"'),
    );
    delete archive["ppt/presentation.xml"];
    delete archive["ppt/_rels/presentation.xml.rels"];

    const patched = parsePresentationSync(
      await patchPresentation({
        data: zipSync(archive),
        commands: [{ type: "remove", at: { field: "slides", route: [{ index: 1 }] } }],
      }),
    );

    expect(patched.slides).toHaveLength(1);
    expect(patched.customShows?.[0]?.slides).toHaveLength(1);
  });

  it("updates canonical package metadata", async () => {
    const source = await generatePresentation(presentation([slide("A")]));
    const parsed = parsePresentationSync(
      await patchPresentation({
        data: source,
        commands: [{ type: "update", at: { field: "title" }, set: "Patched title" }],
      }),
    );
    expect(parsed.title).toBe("Patched title");
  });

  it("inserts slide comments into an existing collection", async () => {
    const added: SlideCommentOptions[] = [{ author: "Bob", text: "added", x: 10, y: 10 }];
    const withComments = presentation([
      {
        ...slide("A"),
        comments: [{ author: "Alice", text: "original", x: 0, y: 0 }],
      },
    ]);
    const source = await generatePresentation(withComments);
    const parsed = parsePresentationSync(
      await patchPresentation({
        data: source,
        commands: [
          {
            type: "insert",
            at: { field: "slides", route: [{ index: 0 }, { collection: "comments" }] },
            value: added,
          },
        ],
      }),
    );
    expect(parsed.slides?.[0]?.comments).toHaveLength(2);
  });

  it("removes slide-owned package parts and identities for replacement slides", async () => {
    const source = await generatePresentation(
      presentation([{ ...slide("A"), notes: "original notes" }, slide("B")]),
    );
    const patched = await patchPresentation({
      data: source,
      commands: [
        { type: "remove", at: { field: "slides", route: [{ index: 0 }] } },
        {
          type: "insert",
          at: { field: "slides" },
          position: "end",
          value: [{ ...slide("C"), notes: "new notes" }],
        },
      ],
    });
    const parsed = parsePresentationSync(patched);
    const { unzipSync } = await import("@office-open/core/opc");
    const archive = unzipSync(patched as Uint8Array);
    const contentTypes = new TextDecoder().decode(archive["[Content_Types].xml"]!);
    const paths = parsed.slides?.map((slide) => slide.sourcePath);

    expect(paths).toEqual(["ppt/slides/slide2.xml", "ppt/slides/slide3.xml"]);
    expect(new Set(paths)).toHaveLength(2);
    expect(new Set(parsed.slides?.map((slide) => slide.slideId))).toHaveLength(2);
    expect(archive["ppt/notesSlides/notesSlide1.xml"]).toBeUndefined();
    expect(archive["ppt/notesSlides/notesSlide2.xml"]).toBeDefined();
    expect(contentTypes).not.toContain('PartName="/ppt/notesSlides/notesSlide1.xml"');
    expect(contentTypes).toContain('PartName="/ppt/notesSlides/notesSlide2.xml"');
  });

  it("reports a missing route with a structured diagnostic", async () => {
    const source = await generatePresentation(presentation([slide("A")]));
    const error = await patchPresentation({
      data: source,
      commands: [
        {
          type: "replace",
          at: { field: "slides", route: [{ name: "Missing" }] },
          value: slide("X"),
        },
      ],
    }).catch((reason: unknown) => reason);

    expect(error).toBeInstanceOf(PatchError);
    expect(error).toMatchObject({
      diagnostic: {
        operation: "replace",
        part: "[canonical-model]",
        reason: "missing-target",
      },
    });
  });
});
