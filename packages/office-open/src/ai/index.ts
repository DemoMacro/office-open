import type { DocumentOptions } from "@office-open/docx";
import type { PresentationOptions } from "@office-open/pptx";
import type { WorkbookOptions } from "@office-open/xlsx";
import { tool } from "ai";

export { formatToolError } from "./error";

import { generate } from "../generate";
import { docxSchema } from "../schemas/docx";
import { pptxSchema } from "../schemas/pptx";
import { xlsxSchema } from "../schemas/xlsx";
import { formatToolError } from "./error";

export const docxTool = tool({
  description:
    "Generate a .docx Word document. " +
    "The input is the document options directly — must include a 'sections' array. " +
    "Each section has 'children' (paragraphs, tables, etc.). " +
    "IMPORTANT: " +
    "Section children must use wrapper keys: { paragraph: {...} }, { table: {...} }, { image: {...} }, etc. " +
    "Paragraph children must use: { text: '...', bold?: true, italic?: true, size?: number, ... }. " +
    "The 'text' key is required in run objects. Plain strings are also accepted. " +
    "Colors are hex WITHOUT '#': 'FF0000', not '#FF0000'. " +
    "Optional metadata: title, creator, subject, styles, numbering, comments, footnotes, endnotes, background, features.",
  inputSchema: docxSchema,
  execute: async (options) => {
    try {
      const base64 = (await generate({
        type: "docx",
        options: options as unknown as DocumentOptions,
        outputType: "base64",
      })) as string;
      return {
        base64,
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      };
    } catch (error) {
      throw new Error(formatToolError("docx", error));
    }
  },
});

export const pptxTool = tool({
  description:
    "Generate a .pptx PowerPoint presentation. " +
    "The input is the document options directly — must include a 'slides' array. " +
    "Each slide has 'children' (shapes, pictures, tables, charts, groups, etc.). " +
    "Shapes use { shape: { x, y, width, height, textBody?, fill?, ... } }. " +
    "IMPORTANT: " +
    "Shape positions (x, y, width, height) are in pixels. " +
    "Colors are hex WITHOUT '#': 'FF0000', not '#FF0000'. " +
    "Fill can be a hex color string or a fill object: '4472C4' or { type: 'solidFill', color: '4472C4' }. " +
    "Optional: size ('16:9' or '4:3' or { width, height }), title, creator, masters, show.",
  inputSchema: pptxSchema,
  execute: async (options) => {
    try {
      const base64 = (await generate({
        type: "pptx",
        options: options as PresentationOptions,
        outputType: "base64",
      })) as string;
      return {
        base64,
        mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      };
    } catch (error) {
      throw new Error(formatToolError("pptx", error));
    }
  },
});

export const xlsxTool = tool({
  description:
    "Generate a .xlsx Excel spreadsheet. " +
    "The input is the document options directly — must include a 'worksheets' array. " +
    "Each worksheet has 'rows' — an array of row objects, each with 'cells'. " +
    "Cell values: string, number, boolean, null. Use 'style' for formatting. " +
    "IMPORTANT: " +
    "Cells can be shorthand values (string, number, boolean) or objects: { value: 'hello', style: { ... } }. " +
    "Column widths use 'width' as a number. " +
    "Optional: columns, mergeCells, freezePanes, autoFilter, images, charts, dataValidations, conditionalFormats.",
  inputSchema: xlsxSchema,
  execute: async (options) => {
    try {
      const base64 = (await generate({
        type: "xlsx",
        options: options as WorkbookOptions,
        outputType: "base64",
      })) as string;
      return {
        base64,
        mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      };
    } catch (error) {
      throw new Error(formatToolError("xlsx", error));
    }
  },
});

export const officeOpenTools = {
  "generate-docx": docxTool,
  "generate-pptx": pptxTool,
  "generate-xlsx": xlsxTool,
} as const;
