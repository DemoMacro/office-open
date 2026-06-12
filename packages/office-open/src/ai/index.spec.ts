import { describe, expect, it } from "vite-plus/test";

import { formatToolError } from "./error";
import { officeOpenTools } from "./index";

describe("officeOpenTools", () => {
  it("should export three tools with correct keys", () => {
    const keys = Object.keys(officeOpenTools);
    expect(keys).toContain("generate-docx");
    expect(keys).toContain("generate-pptx");
    expect(keys).toContain("generate-xlsx");
    expect(keys).toHaveLength(3);
  });

  it("each tool should have a function execute", () => {
    for (const tool of Object.values(officeOpenTools)) {
      expect(typeof tool.execute).toBe("function");
    }
  });

  it("each tool should have a description with IMPORTANT rules", () => {
    for (const tool of Object.values(officeOpenTools)) {
      expect(tool.description).toContain("IMPORTANT:");
    }
  });
});

describe("formatToolError", () => {
  it("should format unsupported paragraph child type errors", () => {
    const msg = formatToolError("docx", new Error("Unsupported paragraph child type: bold"));
    expect(msg).toContain("Invalid paragraph child");
    expect(msg).toContain("bold");
    expect(msg).toContain("{ paragraph:");
    expect(msg).toContain("Do not use raw property names");
  });

  it("should format unsupported run child type errors", () => {
    const msg = formatToolError("docx", new Error("Unsupported run child type: bold, italic"));
    expect(msg).toContain("Invalid run child");
    expect(msg).toContain('"text" key is required');
  });

  it("should format unknown section child type errors", () => {
    const msg = formatToolError("docx", new Error("Unknown section child type"));
    expect(msg).toContain("Unknown section child");
    expect(msg).toContain("{ paragraph:");
    expect(msg).toContain("{ table:");
  });

  it("should format not iterable errors for docx", () => {
    const msg = formatToolError("docx", new Error("options.sections is not iterable"));
    expect(msg).toContain('"sections" must be an array');
  });

  it("should format not iterable errors for pptx", () => {
    const msg = formatToolError("pptx", new Error("options.slides is not iterable"));
    expect(msg).toContain('"slides" must be an array');
  });

  it("should format not iterable errors for xlsx", () => {
    const msg = formatToolError("xlsx", new Error("options.worksheets is not iterable"));
    expect(msg).toContain('"worksheets" must be an array');
  });

  it("should format unknown errors as fallback", () => {
    const msg = formatToolError("docx", new Error("something unexpected happened"));
    expect(msg).toContain("DOCX generation failed");
    expect(msg).toContain("something unexpected happened");
  });

  it("should handle non-Error throws", () => {
    const msg = formatToolError("docx", "string error");
    expect(msg).toContain("DOCX generation failed");
    expect(msg).toContain("string error");
  });
});
