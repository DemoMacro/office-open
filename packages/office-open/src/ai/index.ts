/**
 * Vercel AI SDK tools for generating Office documents.
 *
 * The generate tools use skeleton input schemas (top-level shape + wrapper
 * keys only) instead of the full format schema (~675 KB for docx) — no
 * provider accepts that in a tool definition. Precise field schemas are
 * fetched on demand through the office-open-schema-lookup tool, and the
 * authoritative ajv validation runs inside execute with instancePath-
 * qualified errors the model can iterate on.
 *
 * @module
 */
import type { DocumentOptions } from "@office-open/docx";
import type { PresentationOptions } from "@office-open/pptx";
import type { WorkbookOptions } from "@office-open/xlsx";
import { jsonSchema, tool } from "ai";

export { formatToolError } from "./error";

import { lintWorkbookFormulas } from "@office-open/xlsx";

import { generate } from "../generate";
import { getSkeletonSchema, sliceDocumentSchema, validateDocumentInput } from "../schemas";
import type { DocumentType } from "../schemas/schemas";
import { UnknownDefinitionError } from "../schemas/slice";
import { formatToolError } from "./error";
import { generateVerifiedBase64 } from "./verify";

/** Input accepted by the office-open-schema-lookup tool. */
export interface SchemaLookupInput {
  type: DocumentType;
  definitions: string[];
}

const SKELETON_GUIDANCE =
  "This schema is a skeleton — stubs name the definition they stand for. Fetch real fields with the " +
  "office-open-schema-lookup tool, e.g. { type: 'docx', definitions: ['ParagraphOptions', 'RunOptions'] }. " +
  "Invalid input is rejected with instance-path errors; fix and retry.";

/**
 * The generate tools return the file as base64 for the client UI, but the
 * model only needs the outcome — full base64 in the model context would burn
 * thousands of tokens per document.
 */
function documentGeneratedSummary(output: { base64: string; mimeType: string }): string {
  const kb = Math.ceil((output.base64.length * 3) / 4 / 1024);
  return `Document generated and all validations passed (${output.mimeType}, ${kb} KB).`;
}

export const docxTool = tool({
  description:
    "Generate a .docx Word document. " +
    "The input is the document options directly — must include a 'sections' array. " +
    "Conventions: " +
    "section children are wrapper-key objects ({ paragraph: {...} }, { table: {...} }, …); " +
    "run objects require a 'text' key (plain strings also accepted); " +
    "colors are hex WITHOUT '#': 'FF0000', not '#FF0000'. " +
    SKELETON_GUIDANCE,
  inputSchema: jsonSchema<DocumentOptions>(getSkeletonSchema("docx")),
  execute: async (options) => {
    try {
      const validated = validateDocumentInput("docx", options);
      const bytes = (await generate({
        type: "docx",
        options: validated as unknown as DocumentOptions,
        outputType: "uint8array",
      })) as Uint8Array;
      return {
        base64: generateVerifiedBase64("docx", bytes),
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      };
    } catch (error) {
      throw new Error(formatToolError("docx", error));
    }
  },
  toModelOutput: ({ output }) => ({ type: "text", value: documentGeneratedSummary(output) }),
});

export const pptxTool = tool({
  description:
    "Generate a .pptx PowerPoint presentation. " +
    "The input is the presentation options directly — must include a 'slides' array. " +
    "Conventions: " +
    "shape x/y/width/height take UniversalMeasure strings ('2cm', '1in', '96px') or raw EMU numbers (914400 = 1 inch); " +
    "fills are hex color strings or fill objects ('4472C4' or { type: 'solidFill', color: '4472C4' }); " +
    "colors are hex WITHOUT '#': 'FF0000', not '#FF0000'. " +
    SKELETON_GUIDANCE,
  inputSchema: jsonSchema<PresentationOptions>(getSkeletonSchema("pptx")),
  execute: async (options) => {
    try {
      const validated = validateDocumentInput("pptx", options);
      const bytes = (await generate({
        type: "pptx",
        options: validated as unknown as PresentationOptions,
        outputType: "uint8array",
      })) as Uint8Array;
      return {
        base64: generateVerifiedBase64("pptx", bytes),
        mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      };
    } catch (error) {
      throw new Error(formatToolError("pptx", error));
    }
  },
  toModelOutput: ({ output }) => ({ type: "text", value: documentGeneratedSummary(output) }),
});

export const xlsxTool = tool({
  description:
    "Generate a .xlsx Excel spreadsheet. " +
    "The input is the workbook options directly — must include a 'worksheets' array. " +
    "Conventions: " +
    "cells are shorthand values (string, number, boolean, null) or { value, style } objects; " +
    "column 'width' is in Excel character units. " +
    SKELETON_GUIDANCE,
  inputSchema: jsonSchema<WorkbookOptions>(getSkeletonSchema("xlsx")),
  execute: async (options) => {
    try {
      const validated = validateDocumentInput("xlsx", options);
      const formulaIssues = lintWorkbookFormulas(validated as unknown as WorkbookOptions);
      if (formulaIssues.length > 0) {
        const lines = formulaIssues.map(
          (i) => `  ${i.location}: ${i.message} — formula "${i.formula}"`,
        );
        throw new Error(
          `Invalid xlsx formulas:\n${lines.join("\n")}\nFix the formula or add the missing worksheet.`,
        );
      }
      const bytes = (await generate({
        type: "xlsx",
        options: validated as unknown as WorkbookOptions,
        outputType: "uint8array",
      })) as Uint8Array;
      return {
        base64: generateVerifiedBase64("xlsx", bytes),
        mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      };
    } catch (error) {
      throw new Error(formatToolError("xlsx", error));
    }
  },
  toModelOutput: ({ output }) => ({ type: "text", value: documentGeneratedSummary(output) }),
});

export const schemaLookupTool = tool({
  description:
    "Fetch the precise JSON Schema (draft-07) for office-open option definitions on demand. " +
    "Use it before filling complex objects into the generate tools: the generate input schemas are " +
    "skeletons whose stubs name the definition to look up here. " +
    "Valid names come from the skeleton stubs, or list indexed entries with " +
    "`npx office-open schema index <type>` (all names with --all). " +
    "Returns the requested definitions plus their dependency closure; cataloged domains not " +
    "requested stay as expandable stubs, so request each domain root you need (e.g. " +
    "['ParagraphOptions', 'RunOptions', 'TableOptions']).",
  inputSchema: jsonSchema<SchemaLookupInput>({
    type: "object",
    additionalProperties: false,
    required: ["type", "definitions"],
    properties: {
      type: {
        type: "string",
        enum: ["docx", "pptx", "xlsx"],
        description: "Document format whose schema to slice",
      },
      definitions: {
        type: "array",
        items: { type: "string" },
        minItems: 1,
        maxItems: 8,
        description:
          "Definition names (TS type names, e.g. ParagraphOptions, SlideOptions, StyleOptions). " +
          "At most 8 per call.",
      },
    },
  }),
  execute: async ({ type, definitions }) => {
    try {
      const slice = sliceDocumentSchema(type, definitions);
      return { type, requested: definitions, definitions: slice.definitions };
    } catch (error) {
      if (error instanceof UnknownDefinitionError) {
        // Data, not a throw: lets the model self-correct from the suggestions.
        return {
          type,
          requested: definitions,
          error:
            `${error.message}. Closest: ${error.suggestions.join(", ") || "none"}. ` +
            `List indexed entries with: npx office-open schema index ${type}`,
          suggestions: error.suggestions,
        };
      }
      throw error;
    }
  },
});

export const officeOpenTools = {
  "generate-docx": docxTool,
  "generate-pptx": pptxTool,
  "generate-xlsx": xlsxTool,
  "office-open-schema-lookup": schemaLookupTool,
} as const;
