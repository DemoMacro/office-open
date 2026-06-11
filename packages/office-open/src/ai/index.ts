import type { DocumentOptions } from "@office-open/docx";
import type { PresentationOptions } from "@office-open/pptx";
import type { WorkbookOptions } from "@office-open/xlsx";
import { tool } from "ai";

import { generate } from "../generate";
import { validateDocumentInput } from "../schemas";
import { docxSchema } from "../schemas/docx";
import { pptxSchema } from "../schemas/pptx";
import { xlsxSchema } from "../schemas/xlsx";

export const docxTool = tool({
  description:
    "Generate a .docx Word document. " +
    "The input is the document options directly — must include a 'sections' array. " +
    "Each section has 'children' (paragraphs, tables, etc.). " +
    "Paragraphs use { paragraph: { children: [{ text: '...' }] }}, tables use { table: { rows: [...] } }. " +
    "Optional metadata: title, creator, subject, styles, numbering, comments, footnotes, endnotes, background, features.",
  inputSchema: docxSchema,
  execute: async (options) => {
    const validated = validateDocumentInput("docx", options);
    const base64 = (await generate({
      type: "docx",
      options: validated as unknown as DocumentOptions,
      outputType: "base64",
    })) as string;
    return {
      base64,
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    };
  },
});

export const pptxTool = tool({
  description:
    "Generate a .pptx PowerPoint presentation. " +
    "The input is the document options directly — must include a 'slides' array. " +
    "Each slide has 'children' (shapes, pictures, tables, charts, groups, etc.). " +
    "Shapes use { shape: { x, y, width, height, text?, fill?, ... } }. " +
    "Optional: size ('16:9' or '4:3' or { width, height }), title, creator, masters, show.",
  inputSchema: pptxSchema,
  execute: async (options) => {
    const validated = validateDocumentInput("pptx", options);
    const base64 = (await generate({
      type: "pptx",
      options: validated as PresentationOptions,
      outputType: "base64",
    })) as string;
    return {
      base64,
      mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    };
  },
});

export const xlsxTool = tool({
  description:
    "Generate a .xlsx Excel spreadsheet. " +
    "The input is the document options directly — must include a 'worksheets' array. " +
    "Each worksheet has 'rows' — an array of row objects, each with 'cells'. " +
    "Cell values: string, number, boolean. Use 'style' for formatting. " +
    "Optional: columns, mergeCells, freezePanes, autoFilter, images, charts, dataValidations, conditionalFormats.",
  inputSchema: xlsxSchema,
  execute: async (options) => {
    const validated = validateDocumentInput("xlsx", options);
    const base64 = (await generate({
      type: "xlsx",
      options: validated as WorkbookOptions,
      outputType: "base64",
    })) as string;
    return {
      base64,
      mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    };
  },
});

export const officeOpenTools = {
  "generate-docx": docxTool,
  "generate-pptx": pptxTool,
  "generate-xlsx": xlsxTool,
} as const;
