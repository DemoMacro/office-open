import { describe, expect, it } from "vite-plus/test";

import { FIELD_SPECS } from "../src/descriptor/field-spec";
import { extractInterfaceFields } from "./interface-fields";

const DOCX = "packages/docx/src/parts";

/** Maps a FIELD_SPECS id to the interface + source file that declares it. */
const INTERFACE_SOURCE: Record<string, { interfaceName: string; file: string }> = {
  "core-properties": {
    interfaceName: "CorePropertiesOptions",
    file: "packages/core/src/opc/core.ts",
  },
  "paragraph-properties": {
    interfaceName: "ParagraphPropertiesOptions",
    file: `${DOCX}/paragraph/properties.ts`,
  },
};

describe("extractInterfaceFields", () => {
  it("reads a flat interface (CorePropertiesOptions, 15 fields, sorted)", () => {
    expect(
      extractInterfaceFields("CorePropertiesOptions", "packages/core/src/opc/core.ts"),
    ).toEqual([
      "category",
      "contentStatus",
      "created",
      "creator",
      "description",
      "identifier",
      "keywords",
      "language",
      "lastModifiedBy",
      "lastPrinted",
      "modified",
      "revision",
      "subject",
      "title",
      "version",
    ]);
  });

  it("resolves `&` intersection — ParagraphPropertiesOptions merges all arms", () => {
    // ParagraphPropertiesOptions = {...} & ParagraphStylePropertiesOptions
    // & LevelParagraphStylePropertiesOptions. getInterface().getProperties()
    // would return only the direct members; the Type-level read must return
    // the union: 4 (Base) + 3 (Style) + 32 (Level) + 1 (Options) = 40, before
    // sugar exclusion.
    const fields = extractInterfaceFields(
      "ParagraphPropertiesOptions",
      `${DOCX}/paragraph/properties.ts`,
    );
    expect(fields).toHaveLength(40);
    // Sugar that the exclude list removes later:
    expect(fields).toContain("thematicBreak");
    expect(fields).toContain("rightTabStop");
    expect(fields).toContain("leftTabStop");
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
