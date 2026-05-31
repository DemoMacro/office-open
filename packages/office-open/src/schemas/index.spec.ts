import { readFileSync } from "node:fs";
import { join } from "node:path";

import { describe, expect, it } from "vite-plus/test";

import { docxSchema, pptxSchema, xlsxSchema, validateDocumentInput } from "./index";

const demoDir = join(process.cwd(), "demo");
const docxDemo = JSON.parse(readFileSync(join(demoDir, "1-docx.json"), "utf8"));
const pptxDemo = JSON.parse(readFileSync(join(demoDir, "2-pptx.json"), "utf8"));
const xlsxDemo = JSON.parse(readFileSync(join(demoDir, "3-xlsx.json"), "utf8"));

describe("docxSchema", () => {
  it("should accept valid minimal input", () => {
    const result = docxSchema.safeParse({
      sections: [{ children: [{ paragraph: { text: "Hello" } }] }],
    });
    expect(result.success).toBe(true);
  });

  it("should accept optional metadata", () => {
    const result = docxSchema.safeParse({
      title: "Test",
      creator: "Author",
      sections: [{ children: [] }],
    });
    expect(result.success).toBe(true);
  });

  it("should reject missing sections", () => {
    const result = docxSchema.safeParse({});
    expect(result.success).toBe(false);
  });

  it("should reject non-array sections", () => {
    const result = docxSchema.safeParse({ sections: "invalid" });
    expect(result.success).toBe(false);
  });

  it("should accept valid view enum values", () => {
    const views = ["none", "print", "outline", "masterPages", "normal", "web"] as const;
    for (const view of views) {
      expect(docxSchema.safeParse({ sections: [{ children: [] }], view }).success).toBe(true);
    }
  });

  it("should reject invalid view value", () => {
    const result = docxSchema.safeParse({ sections: [{ children: [] }], view: "invalid" });
    expect(result.success).toBe(false);
  });

  it("should accept valid characterSpacingControl values", () => {
    const vals = ["compressPunctuation", "doNotCompress"] as const;
    for (const val of vals) {
      expect(
        docxSchema.safeParse({ sections: [{ children: [] }], characterSpacingControl: val })
          .success,
      ).toBe(true);
    }
  });

  it("should reject invalid characterSpacingControl", () => {
    const result = docxSchema.safeParse({
      sections: [{ children: [] }],
      characterSpacingControl: "invalid",
    });
    expect(result.success).toBe(false);
  });

  it("should accept valid zoom structure", () => {
    const result = docxSchema.safeParse({
      sections: [{ children: [] }],
      zoom: { percent: 150, val: "fullPage" },
    });
    expect(result.success).toBe(true);
  });

  it("should reject invalid zoom.val", () => {
    const result = docxSchema.safeParse({
      sections: [{ children: [] }],
      zoom: { val: "invalid" },
    });
    expect(result.success).toBe(false);
  });

  it("should accept valid docVars", () => {
    const result = docxSchema.safeParse({
      sections: [{ children: [] }],
      docVars: [{ name: "var1", val: "value1" }],
    });
    expect(result.success).toBe(true);
  });

  it("should reject docVars without name", () => {
    const result = docxSchema.safeParse({
      sections: [{ children: [] }],
      docVars: [{ val: "value1" }],
    });
    expect(result.success).toBe(false);
  });

  it("should accept headers/footers with known keys", () => {
    const result = docxSchema.safeParse({
      sections: [
        {
          children: [],
          headers: { default: { children: [] } },
          footers: { first: { children: [] } },
        },
      ],
    });
    expect(result.success).toBe(true);
  });

  it("should accept section child discriminants", () => {
    for (const child of [
      { paragraph: { text: "Hello" } },
      { table: { rows: [{ cells: [{ children: [{ paragraph: "A" }] }] }] } },
      { toc: { alias: "TOC" } },
      { textbox: { children: [] } },
    ]) {
      expect(docxSchema.safeParse({ sections: [{ children: [child] }] }).success).toBe(true);
    }
  });
});

describe("pptxSchema", () => {
  it("should accept valid minimal input", () => {
    const result = pptxSchema.safeParse({
      slides: [{ children: [] }],
    });
    expect(result.success).toBe(true);
  });

  it("should accept size as string", () => {
    const result = pptxSchema.safeParse({
      size: "16:9",
      slides: [{ children: [] }],
    });
    expect(result.success).toBe(true);
  });

  it("should accept size as object", () => {
    const result = pptxSchema.safeParse({
      size: { width: 1280, height: 720 },
      slides: [{ children: [] }],
    });
    expect(result.success).toBe(true);
  });

  it("should reject missing slides", () => {
    const result = pptxSchema.safeParse({});
    expect(result.success).toBe(false);
  });

  it("should reject invalid size", () => {
    const result = pptxSchema.safeParse({
      size: "A3",
      slides: [{ children: [] }],
    });
    expect(result.success).toBe(false);
  });

  it("should accept valid layout enum values", () => {
    const validLayouts = [
      "blank",
      "title",
      "tx",
      "twoColTx",
      "titleOnly",
      "obj",
      "secHead",
      "chart",
      "tbl",
      "picTx",
      "twoObj",
      "twoTxTwoObj",
      "objTx",
      "vertTx",
      "vertTitleAndTx",
    ] as const;
    for (const layout of validLayouts) {
      const r = pptxSchema.safeParse({ slides: [{ children: [], layout }] });
      expect(r.success).toBe(true);
    }
  });

  it("should reject invalid layout value", () => {
    const result = pptxSchema.safeParse({ slides: [{ children: [], layout: "invalid" }] });
    expect(result.success).toBe(false);
  });

  it("should accept valid show structure", () => {
    const result = pptxSchema.safeParse({
      slides: [{ children: [] }],
      show: { loop: true, kiosk: false, showNarration: true, useTimings: false },
    });
    expect(result.success).toBe(true);
  });

  it("should reject show with non-boolean values", () => {
    const result = pptxSchema.safeParse({
      slides: [{ children: [] }],
      show: { loop: "yes" },
    });
    expect(result.success).toBe(false);
  });

  it("should accept slide comments with required fields", () => {
    const result = pptxSchema.safeParse({
      slides: [{ children: [], comments: [{ author: "John", text: "Test", x: 100, y: 200 }] }],
    });
    expect(result.success).toBe(true);
  });

  it("should reject comments missing required fields", () => {
    const result = pptxSchema.safeParse({
      slides: [{ children: [], comments: [{ author: "John" }] }],
    });
    expect(result.success).toBe(false);
  });

  it("should accept slide child discriminants", () => {
    for (const child of [
      { shape: { x: 100, y: 100, width: 200, height: 100 } },
      { picture: { x: 0, y: 0, width: 100, height: 100 } },
      { table: { rows: [{ cells: [{ children: [] }] }] } },
      { chart: { x: 0, y: 0, width: 100, height: 100 } },
      { line: { x1: 0, y1: 0, x2: 100, y2: 100 } },
      { connector: { x1: 0, y1: 0, x2: 100, y2: 100 } },
      { video: { x: 0, y: 0, width: 100, height: 100 } },
      { audio: { x: 0, y: 0, width: 100, height: 100 } },
      { group: { x: 0, y: 0, width: 100, height: 100, children: [] } },
      { smartart: { x: 0, y: 0, width: 100, height: 100, nodes: [] } },
    ]) {
      expect(pptxSchema.safeParse({ slides: [{ children: [child] }] }).success).toBe(true);
    }
  });
});

describe("xlsxSchema", () => {
  it("should accept valid minimal input", () => {
    const result = xlsxSchema.safeParse({
      worksheets: [{ rows: [] }],
    });
    expect(result.success).toBe(true);
  });

  it("should accept worksheets with cells", () => {
    const result = xlsxSchema.safeParse({
      worksheets: [{ rows: [{ cells: [{ value: "Hello" }, { value: 42 }] }] }],
    });
    expect(result.success).toBe(true);
  });

  it("should accept boolean and null cell values", () => {
    const result = xlsxSchema.safeParse({
      worksheets: [{ rows: [{ cells: [{ value: true }, { value: null }] }] }],
    });
    expect(result.success).toBe(true);
  });

  it("should accept primitive cell shorthand", () => {
    const result = xlsxSchema.safeParse({
      worksheets: [{ rows: [{ cells: ["text", 42, true] }] }],
    });
    expect(result.success).toBe(true);
  });

  it("should accept cell objects with style and formula", () => {
    const result = xlsxSchema.safeParse({
      worksheets: [
        {
          rows: [
            {
              cells: [
                {
                  value: "Hello",
                  style: { font: { bold: true } },
                  formula: { formula: "SUM(A1:A10)" },
                },
              ],
            },
          ],
        },
      ],
    });
    expect(result.success).toBe(true);
  });

  it("should reject missing worksheets", () => {
    const result = xlsxSchema.safeParse({});
    expect(result.success).toBe(false);
  });

  it("should accept column options with required min/max", () => {
    const result = xlsxSchema.safeParse({
      worksheets: [{ columns: [{ min: 1, max: 3, width: 20 }] }],
    });
    expect(result.success).toBe(true);
  });

  it("should reject columns missing required min", () => {
    const result = xlsxSchema.safeParse({
      worksheets: [{ columns: [{ max: 3 }] }],
    });
    expect(result.success).toBe(false);
  });

  it("should reject columns with non-number min", () => {
    const result = xlsxSchema.safeParse({
      worksheets: [{ columns: [{ min: "A", max: "C" }] }],
    });
    expect(result.success).toBe(false);
  });

  it("should accept mergeCells with from/to", () => {
    const result = xlsxSchema.safeParse({
      worksheets: [{ mergeCells: [{ from: { row: 1, col: 1 }, to: { row: 3, col: 3 } }] }],
    });
    expect(result.success).toBe(true);
  });

  it("should reject mergeCells missing from/to", () => {
    const result = xlsxSchema.safeParse({
      worksheets: [{ mergeCells: [{ from: { row: 1 } }] }],
    });
    expect(result.success).toBe(false);
  });

  it("should accept freezePanes with row/col", () => {
    const result = xlsxSchema.safeParse({
      worksheets: [{ freezePanes: { row: 1, col: 2 } }],
    });
    expect(result.success).toBe(true);
  });

  it("should reject freezePanes with non-number values", () => {
    const result = xlsxSchema.safeParse({
      worksheets: [{ freezePanes: { row: "top" } }],
    });
    expect(result.success).toBe(false);
  });

  it("should accept dataValidations with required sqref", () => {
    const result = xlsxSchema.safeParse({
      worksheets: [{ dataValidations: [{ sqref: "A1:A10", type: "list", formula1: '"a,b,c"' }] }],
    });
    expect(result.success).toBe(true);
  });

  it("should reject dataValidations missing sqref", () => {
    const result = xlsxSchema.safeParse({
      worksheets: [{ dataValidations: [{ type: "list" }] }],
    });
    expect(result.success).toBe(false);
  });

  it("should reject dataValidations with invalid type", () => {
    const result = xlsxSchema.safeParse({
      worksheets: [{ dataValidations: [{ sqref: "A1:A10", type: "invalid" }] }],
    });
    expect(result.success).toBe(false);
  });

  it("should accept conditionalFormats with rules", () => {
    const result = xlsxSchema.safeParse({
      worksheets: [
        {
          conditionalFormats: [
            {
              sqref: "A1:A10",
              rules: [{ type: "cellIs", operator: "greaterThan", formulas: ["0"] }],
            },
          ],
        },
      ],
    });
    expect(result.success).toBe(true);
  });

  it("should reject conditionalFormats missing rules", () => {
    const result = xlsxSchema.safeParse({
      worksheets: [{ conditionalFormats: [{ sqref: "A1:A10" }] }],
    });
    expect(result.success).toBe(false);
  });

  it("should reject conditionalFormats with invalid rule type", () => {
    const result = xlsxSchema.safeParse({
      worksheets: [{ conditionalFormats: [{ sqref: "A1:A10", rules: [{ type: "invalid" }] }] }],
    });
    expect(result.success).toBe(false);
  });
});

describe("validateDocumentInput", () => {
  it("should return parsed data for valid docx input", () => {
    const input = { sections: [{ children: [] }] };
    const result = validateDocumentInput("docx", input);
    expect(result).toEqual(input);
  });

  it("should return parsed data for valid pptx input", () => {
    const input = { slides: [{ children: [] }] };
    const result = validateDocumentInput("pptx", input);
    expect(result).toEqual(input);
  });

  it("should return parsed data for valid xlsx input", () => {
    const input = { worksheets: [{ rows: [] }] };
    const result = validateDocumentInput("xlsx", input);
    expect(result).toEqual(input);
  });

  it("should throw for invalid docx input", () => {
    expect(() => validateDocumentInput("docx", {})).toThrow("Invalid docx options");
  });

  it("should throw for invalid pptx input", () => {
    expect(() => validateDocumentInput("pptx", {})).toThrow("Invalid pptx options");
  });

  it("should throw for invalid xlsx input", () => {
    expect(() => validateDocumentInput("xlsx", {})).toThrow("Invalid xlsx options");
  });

  it("should include path in error message for nested errors", () => {
    expect(() =>
      validateDocumentInput("pptx", {
        slides: [{ children: "not-an-array" }],
      }),
    ).toThrow('at "slides.0.children"');
  });

  it("should pass through extra properties via passthrough", () => {
    const input = {
      sections: [{ children: [] }],
      extraField: "preserved",
    };
    const result = validateDocumentInput("docx", input);
    expect((result as Record<string, unknown>).extraField).toBe("preserved");
  });
});

describe("demo JSON files", () => {
  it("1-docx.json should pass schema validation", () => {
    const result = docxSchema.safeParse(docxDemo);
    expect(result.success).toBe(true);
  });

  it("2-pptx.json should pass schema validation", () => {
    const result = pptxSchema.safeParse(pptxDemo);
    expect(result.success).toBe(true);
  });

  it("3-xlsx.json should pass schema validation", () => {
    const result = xlsxSchema.safeParse(xlsxDemo);
    expect(result.success).toBe(true);
  });
});
