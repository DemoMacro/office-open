import type { DocumentOptions } from "@office-open/docx";
import type { PresentationOptions } from "@office-open/pptx";
import type { WorkbookOptions } from "@office-open/xlsx";

import { parseDocument, type DocumentParseFormat, type DocumentParseOptions } from "./document";
import { detectOffice } from "./formats";
import {
  parsePresentation,
  type PresentationParseFormat,
  type PresentationParseOptions,
} from "./presentation";
import { parseWorkbook, type WorkbookParseFormat, type WorkbookParseOptions } from "./workbook";

export type OfficeParseOptions = DocumentParseOptions &
  WorkbookParseOptions &
  PresentationParseOptions;

export type ParsedOffice =
  | { format: DocumentParseFormat; options: DocumentOptions }
  | { format: WorkbookParseFormat; options: WorkbookOptions }
  | { format: PresentationParseFormat; options: PresentationOptions };

export async function parseOffice(
  input: Uint8Array | string,
  options?: OfficeParseOptions,
): Promise<ParsedOffice> {
  const info = detectOffice(input);
  if (info.family === "document") {
    return {
      format: info.format as DocumentParseFormat,
      options: await parseDocument(input, options),
    };
  }
  if (info.family === "workbook") {
    return {
      format: info.format as WorkbookParseFormat,
      options: await parseWorkbook(input, options),
    };
  }
  if (info.family === "presentation") {
    return {
      format: info.format as PresentationParseFormat,
      options: await parsePresentation(input, options),
    };
  }
  throw new Error("Encrypted OOXML documents are not supported by the unified parser");
}
