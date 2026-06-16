import { describe, expect, it } from "vite-plus/test";

import { FIELD_SPECS } from "../src/descriptor/field-spec";
import { extractInterfaceFields } from "./interface-fields";

const DOCX = "packages/docx/src/parts";

/** Maps a FIELD_SPECS id to the interface + source file that declares it. */
const INTERFACE_SOURCE: Record<string, { interfaceName: string; file: string }> = {
  "core-properties": { interfaceName: "CorePropertiesInput", file: `${DOCX}/core-properties.ts` },
  "paragraph-properties": {
    interfaceName: "ParagraphPropertiesOptions",
    file: `${DOCX}/paragraph/properties.ts`,
  },
};

describe("extractInterfaceFields", () => {
  it("reads a flat interface (CorePropertiesInput, 8 fields, sorted)", () => {
    expect(extractInterfaceFields("CorePropertiesInput", `${DOCX}/core-properties.ts`)).toEqual([
      "creator",
      "description",
      "keywords",
      "lastModifiedBy",
      "lastPrinted",
      "revision",
      "subject",
      "title",
    ]);
  });

  it("resolves `&` intersection — ParagraphPropertiesOptions merges all arms", () => {
    // ParagraphPropertiesOptions = {...} & ParagraphStylePropertiesOptions
    // & LevelParagraphStylePropertiesOptions. getInterface().getProperties()
    // would return only the direct members; the Type-level read must return
    // the union: 4 (Base) + 3 (Style) + 32 (Level) + 2 (Options) = 41, before
    // sugar exclusion.
    const fields = extractInterfaceFields(
      "ParagraphPropertiesOptions",
      `${DOCX}/paragraph/properties.ts`,
    );
    expect(fields).toHaveLength(41);
    // Sugar that the exclude list removes later:
    expect(fields).toContain("thematicBreak");
    expect(fields).toContain("rightTabStop");
    expect(fields).toContain("leftTabStop");
    expect(fields).toContain("includeIfEmpty");
  });
});

describe("interface drift — declared interfaceFields match the live interface", () => {
  // For every FIELD_SPECS entry whose source interface is known, the hand-held
  // interfaceFields must equal extractInterfaceFields(...) minus excludeFields.
  // Adding/removing an Options field without updating FIELD_SPECS fails here.
  for (const spec of FIELD_SPECS) {
    const source = INTERFACE_SOURCE[spec.id];
    if (!source) continue;

    it(`${spec.id}: interfaceFields === interface − excludeFields`, () => {
      const extracted = extractInterfaceFields(source.interfaceName, source.file);
      const exclude = new Set(spec.excludeFields ?? []);
      const expected = extracted.filter((f) => !exclude.has(f)).sort();
      expect([...spec.interfaceFields].sort()).toEqual(expected);
    });
  }
});
