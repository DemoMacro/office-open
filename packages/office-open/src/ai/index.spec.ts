import { unzipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";

import { detectOffice } from "../formats";
import { formatToolError } from "./error";
import {
  docxTool,
  odpTool,
  odsTool,
  officeOpenTools,
  odtTool,
  pptxTool,
  rtfTool,
  schemaLookupTool,
  xlsxTool,
} from "./index";

/** Tool-definition budget: the full docx schema is ~675 KB — the whole point
 *  of skeletons. Ratchet against silent growth; see skeleton.spec.ts for the
 *  industry yardstick (healthy single-tool schema ≈ 200-600 tokens). */
const MAX_INPUT_SCHEMA_BYTES = 28 * 1024;

const VALID_INPUTS = {
  docx: { sections: [{ children: [{ paragraph: { children: ["AI DOCX"] } }] }] },
  pptx: {
    slides: [
      {
        children: [{ shape: { x: 0, y: 0, width: 10, height: 4, textBody: { text: "AI PPTX" } } }],
      },
    ],
  },
  xlsx: { worksheets: [{ rows: [{ cells: [{ value: "AI XLSX" }] }] }] },
} as const;

async function executeTool(
  tool: unknown,
  input: unknown,
): Promise<{ base64: string; mimeType: string }> {
  const execute = (
    tool as {
      execute: (value: unknown) => Promise<{ base64: string; mimeType: string }>;
    }
  ).execute;
  return execute(input);
}

describe("officeOpenTools", () => {
  it("should export eight tools with correct keys", () => {
    const keys = Object.keys(officeOpenTools);
    expect(keys).toContain("generate-docx");
    expect(keys).toContain("generate-pptx");
    expect(keys).toContain("generate-xlsx");
    expect(keys).toContain("generate-odt");
    expect(keys).toContain("generate-rtf");
    expect(keys).toContain("generate-ods");
    expect(keys).toContain("generate-odp");
    expect(keys).toContain("office-open-schema-lookup");
    expect(keys).toHaveLength(8);
  });

  it("each tool should have a function execute", () => {
    for (const tool of Object.values(officeOpenTools)) {
      expect(typeof tool.execute).toBe("function");
    }
  });

  it("each tool should have a description with conventions", () => {
    for (const [name, tool] of Object.entries(officeOpenTools)) {
      if (name === "office-open-schema-lookup") continue;
      expect(tool.description).toContain("Conventions:");
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

describe("rtfTool", () => {
  it("generates RTF through the aggregate writer", async () => {
    const result = await executeTool(rtfTool, {
      sections: [{ children: [{ paragraph: { children: ["AI RTF"] } }] }],
    });
    const source = new TextDecoder().decode(
      Uint8Array.from(atob(result.base64), (character) => character.charCodeAt(0)),
    );
    expect(result.mimeType).toBe("application/rtf");
    expect(source).toContain("AI RTF");
  });
});

describe("schemaLookupTool", () => {
  it("returns the requested definitions with their closure as type text", async () => {
    const execute = schemaLookupTool.execute as unknown as (input: {
      type: "docx" | "pptx" | "xlsx";
      definitions: string[];
    }) => Promise<{ type: string; requested: string[]; typeText: string }>;
    const result = await execute({ type: "docx", definitions: ["ParagraphOptions"] });
    expect(result.requested).toEqual(["ParagraphOptions"]);
    expect(result.typeText).toContain("ParagraphOptions {");
    // closure renders dependencies; cataloged boundaries (and defs demoted by
    // the slice size cap, ParagraphChild being the largest) stay stubs
    expect(result.typeText).toContain("CnfConditionalOptions {");
    expect(result.typeText).toContain("RunOptions // stub");
    expect(result.typeText).toContain("ParagraphChild // stub");
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

  it("generates a clean document through the OPC gate", async () => {
    const execute = docxTool.execute as (
      input: unknown,
    ) => Promise<{ base64: string; mimeType: string }>;
    const result = await execute({
      sections: [{ children: [{ paragraph: { children: ["Hi"] } }] }],
    });
    expect(result.base64.length).toBeGreaterThan(100);
    expect(result.mimeType).toContain("wordprocessingml");
  });

  it("rejects xlsx formulas referencing missing worksheets before generating", async () => {
    const execute = xlsxTool.execute as (input: unknown) => Promise<unknown>;
    await expect(
      execute({
        worksheets: [{ rows: [{ cells: [{ reference: "B1", formula: "Nope!A1" }] }] }],
      }),
    ).rejects.toThrow("Invalid xlsx formulas");
  });
});

describe("generate tool dispatch", () => {
  it.each([
    ["docxTool", docxTool, VALID_INPUTS.docx, "wordprocessingml.document"],
    ["pptxTool", pptxTool, VALID_INPUTS.pptx, "presentationml.presentation"],
    ["xlsxTool", xlsxTool, VALID_INPUTS.xlsx, "spreadsheetml.sheet"],
  ])("generates %s", async (_name, tool, input, mimeType) => {
    const result = await executeTool(tool, input);
    expect(result.base64).toBeTruthy();
    expect(result.mimeType).toBe(`application/vnd.openxmlformats-officedocument.${mimeType}`);
  });

  it("generates ODF through canonical schemas and verifies package structure", async () => {
    const expected = {
      odt: {
        tool: odtTool,
        input: VALID_INPUTS.docx,
        mimeType: "application/vnd.oasis.opendocument.text",
        files: ["content.xml", "styles.xml", "meta.xml"],
      },
      ods: {
        tool: odsTool,
        input: VALID_INPUTS.xlsx,
        mimeType: "application/vnd.oasis.opendocument.spreadsheet",
        files: ["content.xml", "styles.xml", "meta.xml"],
      },
      odp: {
        tool: odpTool,
        input: VALID_INPUTS.pptx,
        mimeType: "application/vnd.oasis.opendocument.presentation",
        files: ["content.xml", "styles.xml", "meta.xml"],
      },
    } as const;

    for (const [format, fixture] of Object.entries(expected)) {
      const result = await executeTool(fixture.tool, fixture.input);
      const bytes = new Uint8Array(Buffer.from(result.base64, "base64"));
      expect(result.mimeType).toBe(fixture.mimeType);
      expect(detectOffice(bytes).format).toBe(format);
      const files = unzipSync(bytes);
      expect(new TextDecoder().decode(files.mimetype).trim()).toBe(fixture.mimeType);
      const manifest = new TextDecoder().decode(files["META-INF/manifest.xml"]);
      expect(manifest).toContain(`manifest:full-path="/"`);
      expect(manifest).toContain(`manifest:media-type="${fixture.mimeType}"`);
      for (const file of fixture.files) expect(files[file]).toBeInstanceOf(Uint8Array);
    }
  });

  it.each([
    ["odtTool", odtTool, "docx"],
    ["odsTool", odsTool, "xlsx"],
    ["odpTool", odpTool, "pptx"],
  ])("rejects invalid %s tool input", async (_name, tool, format) => {
    await expect(executeTool(tool, { invalidRoot: true })).rejects.toThrow(
      `Invalid ${format} options`,
    );
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
