import type { DocumentOptions } from "@parts/core-properties";
import { describe, expect, it } from "vite-plus/test";

import { generateDocument } from "../generate";
import { parseDocumentSync } from "../parse";
import type { SectionOptions } from "../shared/section";
import { patchDocument } from "./patch";

const section = (text: string): SectionOptions => ({
  children: [{ paragraph: { children: [text] } }],
});

const document = (sections: SectionOptions[]): DocumentOptions => ({ sections });

const source = async (): Promise<Uint8Array> =>
  (await generateDocument(document([section("Hello {{name}}"), section("Second")]), {
    type: "uint8array",
  })) as Uint8Array;

describe("patchDocument commands", () => {
  it("replaces all selected text within a section", async () => {
    const parsed = parseDocumentSync(
      await patchDocument({
        data: await source(),
        commands: [
          {
            type: "replace",
            at: { field: "sections", route: [{ index: 0 }], text: { placeholder: "name" } },
            value: "Alice",
          },
        ],
      }),
    );

    expect(JSON.stringify(parsed.sections?.[0])).toContain("Alice");
    expect(JSON.stringify(parsed.sections?.[0])).not.toContain("{{name}}");
    expect(JSON.stringify(parsed.sections?.[1])).toContain("Second");
  });

  it("replaces placeholders split across adjacent runs", async () => {
    const parsed = parseDocumentSync(
      await patchDocument({
        data: await source(),
        commands: [
          {
            type: "replace",
            at: { field: "sections", route: [{ index: 0 }] },
            value: {
              children: [
                {
                  paragraph: { children: [{ text: "Hello {{na" }, { text: "me}}" }] },
                },
              ],
            },
          },
          {
            type: "replace",
            at: { field: "sections", route: [{ index: 0 }], text: { placeholder: "name" } },
            value: "Alice",
          },
        ],
      }),
    );

    expect(JSON.stringify(parsed.sections?.[0])).toContain("Hello Alice");
    expect(JSON.stringify(parsed.sections?.[0])).not.toContain("{{na");
  });

  it("inserts, moves, and removes sections in order", async () => {
    const parsed = parseDocumentSync(
      await patchDocument({
        data: await source(),
        commands: [
          { type: "insert", at: { field: "sections" }, position: "end", value: [section("Third")] },
          { type: "move", at: { field: "sections", route: [{ index: 2 }] }, to: "start" },
          { type: "remove", at: { field: "sections", route: [{ index: 1 }] } },
        ],
      }),
    );

    expect(JSON.stringify(parsed.sections)).toContain("Third");
    expect(parsed.sections).toHaveLength(2);
  });

  it("updates canonical metadata", async () => {
    const parsed = parseDocumentSync(
      await patchDocument({
        data: await source(),
        commands: [{ type: "update", at: { field: "title" }, set: "Patched title" }],
      }),
    );

    expect(parsed.title).toBe("Patched title");
  });

  it("creates a missing settings part from an update command", async () => {
    const parsed = parseDocumentSync(
      await patchDocument({
        data: await source(),
        commands: [{ type: "update", at: { field: "settings" }, set: { updateFields: true } }],
      }),
    );

    expect(parsed.settings?.updateFields).toBe(true);
  });
});
