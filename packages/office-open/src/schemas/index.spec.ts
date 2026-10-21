import { readFileSync } from "node:fs";
import { join } from "node:path";

import { Ajv } from "ajv";
import { describe, expect, it } from "vite-plus/test";

import { docxSchema, pptxSchema, SCHEMAS, validateDocumentInput, xlsxSchema } from "./index";

const demoDir = join(import.meta.dirname, "../../demo");
const docxDemo = JSON.parse(readFileSync(join(demoDir, "1-docx.json"), "utf8"));
const pptxDemo = JSON.parse(readFileSync(join(demoDir, "2-pptx.json"), "utf8"));
const xlsxDemo = JSON.parse(readFileSync(join(demoDir, "3-xlsx.json"), "utf8"));
const schemaDir = join(import.meta.dirname, "../../schemas");
const patchSchemas = Object.fromEntries(
  ["docx", "pptx", "xlsx"].map((type) => [
    type,
    JSON.parse(readFileSync(join(schemaDir, `${type}.patch.schema.json`), "utf8")),
  ]),
);

describe("schema artifacts", () => {
  it("should be draft-07 schemas with a $ref envelope", () => {
    for (const [type, schema] of Object.entries(SCHEMAS)) {
      expect(schema.$schema).toBe("http://json-schema.org/draft-07/schema#");
      expect(schema.$ref).toBe(
        `#/definitions/${type === "docx" ? "DocumentOptions" : type === "pptx" ? "PresentationOptions" : "WorkbookOptions"}`,
      );
      expect(Object.keys(schema.definitions as object).length).toBeGreaterThan(100);
    }
  });
});

describe("patch schema artifacts", () => {
  it("should resolve model refs from their base schema", () => {
    const fields = { docx: "sections", pptx: "slides", xlsx: "worksheets" } as const;

    for (const type of ["docx", "pptx", "xlsx"] as const) {
      const ajv = new Ajv({ allErrors: true, logger: false, strict: false });
      ajv.addSchema(SCHEMAS[type]);
      const validate = ajv.compile(patchSchemas[type]);
      const command = {
        at: { field: fields[type], route: [{ index: 0 }] },
        type: "remove",
      };

      expect(validate(command)).toBe(true);
      expect(validate({ type: "bogus" })).toBe(false);
    }
  }, 30_000);
});

describe("validateDocumentInput", () => {
  it("should accept each format's demo JSON", () => {
    expect(validateDocumentInput("docx", docxDemo)).toEqual(docxDemo);
    expect(validateDocumentInput("pptx", pptxDemo)).toEqual(pptxDemo);
    expect(validateDocumentInput("xlsx", xlsxDemo)).toEqual(xlsxDemo);
  }, 30_000);

  it("should reject input missing the root array", () => {
    // Only DocumentOptions.sections is required; PresentationOptions and
    // WorkbookOptions are all-optional, so mistyped root arrays cover them.
    expect(() => validateDocumentInput("docx", {})).toThrow("Invalid docx options");
    expect(() => validateDocumentInput("pptx", { slides: "not-an-array" })).toThrow(
      "Invalid pptx options",
    );
    expect(() => validateDocumentInput("xlsx", { worksheets: "not-an-array" })).toThrow(
      "Invalid xlsx options",
    );
  });

  it("should include the instance path in nested errors", () => {
    expect(() =>
      validateDocumentInput("pptx", {
        slides: [{ children: "not-an-array" }],
      }),
    ).toThrow('at "slides.0.children"');
  });

  it("should reject values outside a literal union", () => {
    expect(() =>
      validateDocumentInput("xlsx", {
        worksheets: [{ rows: [{ cells: [{ style: { horizontal: "middle" } }] }] }],
      }),
    ).toThrow("Invalid xlsx options");
  });

  it("should reject unknown root properties", () => {
    expect(() =>
      validateDocumentInput("docx", {
        sections: [{ children: [] }],
        extraField: "preserved",
      }),
    ).toThrow("must NOT have additional properties");
  });
});

describe("exported schema objects", () => {
  it("should expose one schema per format", () => {
    expect(docxSchema).toBe(SCHEMAS.docx);
    expect(pptxSchema).toBe(SCHEMAS.pptx);
    expect(xlsxSchema).toBe(SCHEMAS.xlsx);
  });
});
