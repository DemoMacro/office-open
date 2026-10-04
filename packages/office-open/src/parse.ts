import type { DocumentOptions } from "@office-open/docx";
import type { PresentationOptions } from "@office-open/pptx";
import type { WorkbookOptions } from "@office-open/xlsx";

import {
  parseDocumentFile,
  type DocumentFileParseFormat,
  type DocumentFileParseOptions,
} from "./document";
import { detectOfficeFormat } from "./formats";
import {
  parsePresentationFile,
  type PresentationFileParseFormat,
  type PresentationFileParseOptions,
} from "./presentation";
import {
  parseWorkbookFile,
  type WorkbookFileParseFormat,
  type WorkbookFileParseOptions,
} from "./workbook";

type DocumentOfficeFormat = DocumentFileParseFormat;
type WorkbookOfficeFormat = WorkbookFileParseFormat;
type PresentationOfficeFormat = PresentationFileParseFormat;

export type OfficeDocumentParseOptions = DocumentFileParseOptions;
export type OfficeWorkbookParseOptions = WorkbookFileParseOptions;
export type OfficePresentationParseOptions = PresentationFileParseOptions;

export type ParsedOfficeDocument =
  | { type: DocumentOfficeFormat; options: DocumentOptions }
  | { type: WorkbookOfficeFormat; options: WorkbookOptions }
  | { type: PresentationOfficeFormat; options: PresentationOptions };

export async function parseOfficeDocument(
  input: Uint8Array | string,
  options?: DocumentFileParseOptions & WorkbookFileParseOptions & PresentationFileParseOptions,
): Promise<ParsedOfficeDocument> {
  const info = detectOfficeFormat(input);
  if (info.family === "document") {
    return {
      type: info.format as DocumentOfficeFormat,
      options: await parseDocumentFile(input, options),
    };
  }
  if (info.family === "workbook") {
    return {
      type: info.format as WorkbookOfficeFormat,
      options: await parseWorkbookFile(input, options),
    };
  }
  if (info.family === "presentation") {
    return {
      type: info.format as PresentationOfficeFormat,
      options: await parsePresentationFile(input, options),
    };
  }
  throw new Error("Encrypted OOXML documents are not supported by the unified parser");
}
