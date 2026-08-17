import { describe, expect, it } from "vite-plus/test";

import { formatToolError } from "./error";
import { docxTool, officeOpenTools, schemaLookupTool } from "./index";

/** Tool-definition budget: the full docx schema is ~675 KB — the whole point of skeletons. */
const MAX_INPUT_SCHEMA_BYTES = 22 * 1024;

describe("officeOpenTools", () => {
  it("should export four tools with correct keys", () => {
    const keys = Object.keys(officeOpenTools);
    expect(keys).toContain("generate-docx");
    expect(keys).toContain("generate-pptx");
    expect(keys).toContain("generate-xlsx");
    expect(keys).toContain("office-open-schema-lookup");
    expect(keys).toHaveLength(4);
  });

  it("each tool should have a function execute", () => {
    for (const tool of Object.values(officeOpenTools)) {
      expect(typeof tool.execute).toBe("function");
    }
  });

  it("each tool should have a description with IMPORTANT rules", () => {
    for (const [name, tool] of Object.entries(officeOpenTools)) {
      if (name === "office-open-schema-lookup") continue;
      expect(tool.description).toContain("IMPORTANT:");
    }
  });

  it("generate tools stay within the skeleton size budget", () => {
    for (const [name, tool] of Object.entries(officeOpenTools)) {
      if (name === "office-open-schema-lookup") continue;
      const bytes = JSON.stringify(tool.inputSchema ?? {}).length;
      expect(bytes).toBeLessThan(MAX_INPUT_SCHEMA_BYTES);
    }
  });

  it("generate tools point stubs at the lookup tool", () => {
    for (const [name, tool] of Object.entries(officeOpenTools)) {
      if (name === "office-open-schema-lookup") continue;
      expect(tool.description).toContain("office-open-schema-lookup");
    }
  });
});

describe("schemaLookupTool", () => {
  it("returns the requested definitions with their closure", async () => {
    const execute = schemaLookupTool.execute as (input: {
      type: "docx" | "pptx" | "xlsx";
      definitions: string[];
    }) => Promise<{ type: string; requested: string[]; definitions: Record<string, unknown> }>;
    const result = await execute({ type: "docx", definitions: ["ParagraphOptions"] });
    expect(result.requested).toEqual(["ParagraphOptions"]);
    expect(result.definitions.ParagraphOptions).toBeDefined();
    expect(result.definitions.RunOptions).toBeDefined();
  });

  it("returns suggestions (not a throw) for unknown names", async () => {
    const execute = schemaLookupTool.execute as unknown as (input: {
      type: "docx" | "pptx" | "xlsx";
      definitions: string[];
    }) => Promise<{ error?: string; suggestions?: string[] }>;
    const result = await execute({ type: "docx", definitions: ["ParagraphOption"] });
    expect(result.error).toContain("Unknown definition");
    expect(result.suggestions).toContain("ParagraphOptions");
  });
});

describe("generate tool validation gate", () => {
  it("rejects invalid input with an aggregated instance-path error before generating", async () => {
    const execute = docxTool.execute as (
      input: unknown,
    ) => Promise<{ base64: string; mimeType: string }>;
    await expect(execute({ sections: "oops" })).rejects.toThrow("Invalid docx options");
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
