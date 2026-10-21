import { PatchError } from "@office-open/core";
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
