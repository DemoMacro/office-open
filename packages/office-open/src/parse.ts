import { toUint8ArrayAsync, type DataType } from "@office-open/core";
import type { DocumentOptions } from "@office-open/docx";
import type { OdsWorkbookOptions } from "@office-open/ods";
import type { OdtDocumentOptions } from "@office-open/odt";
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
  | {
      format: Exclude<DocumentParseFormat, "odt">;
      options: DocumentOptions;
    }
  | { format: "odt"; options: OdtDocumentOptions }
  | {
      format: Exclude<WorkbookParseFormat, "ods">;
      options: WorkbookOptions;
    }
  | { format: "ods"; options: OdsWorkbookOptions }
  | {
      format: Exclude<PresentationParseFormat, "odp">;
      options: PresentationOptions;
    }
  | { format: "odp"; options: PresentationOptions };

export async function parseOffice(
  input: DataType,
  options?: OfficeParseOptions,
): Promise<ParsedOffice> {
  const data = await toUint8ArrayAsync(input);
  const info = detectOffice(data);
  if (info.family === "document") {
    return {
      format: info.format as DocumentParseFormat,
      options: await parseDocument(data, options),
    } as ParsedOffice;
  }
  if (info.family === "workbook") {
    return {
      format: info.format as WorkbookParseFormat,
      options: await parseWorkbook(data, options),
    } as ParsedOffice;
  }
  if (info.family === "presentation") {
    return {
      format: info.format as PresentationParseFormat,
      options: await parsePresentation(data, options),
    } as ParsedOffice;
  }
  throw new Error("Encrypted OOXML documents are not supported by the unified parser");
}
