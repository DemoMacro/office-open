import { describe, expect, it } from "vite-plus/test";

import { docxSchema, pptxSchema, xlsxSchema, validateDocumentInput } from "./index";

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
});

describe("xlsxSchema", () => {
  it("should accept valid minimal input", () => {
    const result = xlsxSchema.safeParse({
      worksheets: [{ children: [] }],
    });
    expect(result.success).toBe(true);
  });

  it("should accept worksheets with cells", () => {
    const result = xlsxSchema.safeParse({
      worksheets: [{ children: [{ cells: [{ value: "Hello" }, { value: 42 }] }] }],
    });
    expect(result.success).toBe(true);
  });

  it("should accept boolean cell values", () => {
    const result = xlsxSchema.safeParse({
      worksheets: [{ children: [{ cells: [{ value: true }] }] }],
    });
    expect(result.success).toBe(true);
  });

  it("should reject missing worksheets", () => {
    const result = xlsxSchema.safeParse({});
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
    const input = { worksheets: [{ children: [] }] };
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
