import Ajv from "ajv";
import addFormats from "ajv-formats";
import { describe, expect, it } from "vite-plus/test";

import { SCHEMA_ENTRIES } from "./entries";
import { SCHEMAS } from "./index";
import {
  UnknownDefinitionError,
  assertKnownDefinitions,
  sliceDocumentSchema,
  sliceSchema,
} from "./slice";

/** Every internal $ref inside `node` resolves (after percent-decoding) inside `definitions`. */
function assertNoDanglingRefs(node: unknown, definitions: Record<string, unknown>): void {
  const refs = JSON.stringify(node)?.match(/#\/definitions\/([A-Za-z0-9_%.$-]+)/g) ?? [];
  for (const ref of refs) {
    const name = decodeURIComponent(ref.slice(14));
    expect(definitions).toHaveProperty(name);
  }
}

describe("SCHEMA_ENTRIES catalog validity", () => {
  it("every cataloged entry exists in its format's schema definitions", () => {
    for (const [type, entries] of Object.entries(SCHEMA_ENTRIES)) {
      const definitions = SCHEMAS[type as keyof typeof SCHEMAS].definitions as Record<
        string,
        unknown
      >;
      for (const entry of entries) {
        expect(definitions).toHaveProperty(entry.name);
      }
    }
  });
});

describe("sliceSchema", () => {
  it("expands requested entries and stubs cataloged boundaries", () => {
    const slice = sliceDocumentSchema("docx", ["ParagraphOptions"]);
    const definitions = slice.definitions as Record<string, Record<string, unknown>>;

    expect(definitions.ParagraphOptions?.properties).toBeDefined();
    for (const stubbed of [
      "PictureOptions",
      "ChartOptions",
      "SmartArtOptions",
      "GroupOptions",
      "ShapeOptions",
      "SdtRunOptions",
    ]) {
      expect(definitions[stubbed]?.$comment).toBe(`office-open-stub:${stubbed}`);
    }
    // Non-cataloged dependencies expand normally (runs enter via the
    // ParagraphRunOptions entry, which is stubbed — but direct helpers expand).
    expect(definitions.SpacingProperties).toBeDefined();
  });

  it("never stubs a requested entry", () => {
    const slice = sliceDocumentSchema("docx", ["ParagraphOptions", "ChartOptions"]);
    const definitions = slice.definitions as Record<string, Record<string, unknown>>;
    expect(definitions.ChartOptions?.properties).toBeDefined();
    expect(definitions.ChartOptions?.$comment).toBeUndefined();
  });

  it("produces a standalone schema with no dangling refs (circular refs terminate)", () => {
    const slice = sliceDocumentSchema("docx", ["SectionOptions"]);
    assertNoDanglingRefs(slice, slice.definitions as Record<string, unknown>);
  });

  it("wires a single requested definition as the root $ref", () => {
    const slice = sliceDocumentSchema("docx", ["RunOptions"]);
    expect(slice.$ref).toBe("#/definitions/RunOptions");
  });

  it("does not mutate the source schema", () => {
    const before = JSON.stringify(SCHEMAS.docx);
    sliceDocumentSchema("docx", ["ParagraphOptions", "TableOptions"]);
    expect(JSON.stringify(SCHEMAS.docx)).toBe(before);
  });

  it("throws UnknownDefinitionError with suggestions for a misspelled name", () => {
    let caught: unknown;
    try {
      sliceDocumentSchema("docx", ["ParagraphOption"]);
    } catch (error) {
      caught = error;
    }
    expect(caught).toBeInstanceOf(UnknownDefinitionError);
    const err = caught as UnknownDefinitionError;
    expect(err.suggestions).toContain("ParagraphOptions");
    expect(() => assertKnownDefinitions("docx", ["ParagraphOption"])).toThrow(
      UnknownDefinitionError,
    );
  });

  it("supports the percent-encoded generic definition name", () => {
    const slice = sliceDocumentSchema("docx", ["HeaderFooterGroup<HeaderFooterReference>"]);
    const definitions = slice.definitions as Record<string, unknown>;
    expect(definitions).toHaveProperty("HeaderFooterGroup<HeaderFooterReference>");
  });

  it("compiles under ajv across all three formats", () => {
    const ajv = new Ajv({ allErrors: true, allowUnionTypes: true });
    addFormats(ajv);
    const slices = [
      sliceDocumentSchema("docx", ["ParagraphOptions", "RunOptions"]),
      sliceDocumentSchema("pptx", ["SlideOptions", "ShapeOptions"]),
      sliceDocumentSchema("xlsx", ["WorksheetOptions", "RowOptions", "CellOptions"]),
    ];
    for (const slice of slices) {
      expect(() => ajv.compile(slice as never)).not.toThrow();
      assertNoDanglingRefs(slice, slice.definitions as Record<string, unknown>);
    }
  });

  it("slices a schema without a format (no stub boundaries, pure closure)", () => {
    const slice = sliceSchema(SCHEMAS.docx, ["RunStylePropertiesOptions"]);
    const definitions = slice.definitions as Record<string, unknown>;
    for (const name of Object.keys(definitions)) {
      expect((definitions[name] as Record<string, unknown>).$comment).toBeUndefined();
    }
  });
});
