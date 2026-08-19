import { readFileSync } from "node:fs";
import { join } from "node:path";

import Ajv from "ajv";
import addFormats from "ajv-formats";
import { describe, expect, it } from "vite-plus/test";

import type { DocumentType } from "./schemas";
import { getSkeletonSchema } from "./skeleton";

const demoDir = join(import.meta.dirname, "../../demo");
const demos: Record<DocumentType, unknown> = {
  docx: JSON.parse(readFileSync(join(demoDir, "1-docx.json"), "utf8")),
  pptx: JSON.parse(readFileSync(join(demoDir, "2-pptx.json"), "utf8")),
  xlsx: JSON.parse(readFileSync(join(demoDir, "3-xlsx.json"), "utf8")),
};

/** Skeleton budget: must stay far below the full schema (~675 KB for docx).
 *  The budget is a ratchet against silent growth — every field added to the
 *  root options lands here, and the schema rides along in every agent
 *  conversation. Industry measurements put a healthy single-tool schema at
 *  200-600 tokens (~1-3 KB); the current ~25 KB (~6K tokens) is already a
 *  heavyweight tool by that yardstick, so treat this ceiling as damage
 *  control (shrinking the skeleton generator is backlog), not a target. */
const MAX_SKELETON_BYTES = 26 * 1024;

describe("getSkeletonSchema", () => {
  it("compiles under ajv for all three formats", () => {
    const ajv = new Ajv({ allErrors: true, allowUnionTypes: true });
    addFormats(ajv);
    for (const type of ["docx", "pptx", "xlsx"] as const) {
      expect(() => ajv.compile(getSkeletonSchema(type) as never)).not.toThrow();
    }
  });

  it("stays within the skeleton size budget", () => {
    for (const type of ["docx", "pptx", "xlsx"] as const) {
      const bytes = JSON.stringify(getSkeletonSchema(type)).length;
      expect(bytes).toBeLessThan(MAX_SKELETON_BYTES);
    }
  });

  it("memoizes per format", () => {
    expect(getSkeletonSchema("docx")).toBe(getSkeletonSchema("docx"));
  });

  it("expands the docx spine with wrapper keys and stubs everything deeper", () => {
    const skeleton = getSkeletonSchema("docx") as {
      properties: Record<
        string,
        { items?: { properties: Record<string, { items: { anyOf: { title: string }[] } }> } }
      >;
      required?: string[];
    };
    expect(skeleton.required).toContain("sections");
    const section = skeleton.properties.sections?.items;
    expect(section?.properties).toHaveProperty("children");

    const wrapperTitles = section?.properties.children?.items?.anyOf?.map((b) => b.title) ?? [];
    for (const key of ["paragraph", "table", "toc", "textbox"]) {
      expect(wrapperTitles).toContain(key);
    }

    // The paragraph branch's value is { anyOf: [string shorthand, ParagraphOptions stub] }.
    const paragraphBranch = section?.properties.children?.items?.anyOf?.find(
      (b) => b.title === "paragraph",
    ) as { properties: { paragraph: { anyOf?: { $comment?: string }[] } } } | undefined;
    const stubLegs =
      paragraphBranch?.properties.paragraph.anyOf?.filter((leg) => leg.$comment) ?? [];
    expect(stubLegs[0]?.$comment).toBe("office-open-stub:ParagraphOptions");
  });

  it("never rejects real input: each format's demo JSON passes the skeleton", () => {
    const ajv = new Ajv({ allErrors: true, allowUnionTypes: true });
    addFormats(ajv);
    for (const type of ["docx", "pptx", "xlsx"] as const) {
      const validate = ajv.compile(getSkeletonSchema(type) as never);
      const result = validate(demos[type]);
      expect(validate.errors ?? []).toEqual([]);
      expect(result).toBe(true);
    }
  });
});
