import type { OutputByType, OutputType, ReproducibleGenerationOptions } from "@office-open/core";
import type { DocumentOptions } from "@office-open/docx";
import type { OdpPresentationOptions } from "@office-open/odp";
import type { OdsWorkbookOptions } from "@office-open/ods";
import type { OdtDocumentOptions } from "@office-open/odt";
import type { PresentationOptions } from "@office-open/pptx";
import type { WorkbookOptions } from "@office-open/xlsx";

import { generateDocument, type DocumentGenerateFormat } from "./document";
import { generatePresentation, type PresentationGenerateFormat } from "./presentation";
import { generateWorkbook, type WorkbookGenerateFormat } from "./workbook";

export type OfficeGenerateFormat =
  | DocumentGenerateFormat
  | WorkbookGenerateFormat
  | PresentationGenerateFormat;

/** Map from a writable office format to its canonical options model. */
export interface OfficeOptionsFor {
  docx: DocumentOptions;
  docm: DocumentOptions;
  dotx: DocumentOptions;
  dotm: DocumentOptions;
  odt: OdtDocumentOptions;
  rtf: DocumentOptions;
  xlsx: WorkbookOptions;
  xlsm: WorkbookOptions;
  xltx: WorkbookOptions;
  xltm: WorkbookOptions;
  ods: OdsWorkbookOptions;
  pptx: PresentationOptions;
  pptm: PresentationOptions;
  potx: PresentationOptions;
  potm: PresentationOptions;
  odp: OdpPresentationOptions;
}

type OdfGenerateFormat = "odt" | "ods" | "odp";

/** ODF writers are deterministic today but do not support OPC reproducibility controls. */
export type OfficeGenerationOptions<Format extends OfficeGenerateFormat> = {
  reproducible?: Format extends OdfGenerateFormat ? never : ReproducibleGenerationOptions;
};

const DOCUMENT_FORMATS = [
  "docx",
  "docm",
  "dotx",
  "dotm",
  "odt",
  "rtf",
] as const satisfies readonly DocumentGenerateFormat[];
const WORKBOOK_FORMATS = [
  "xlsx",
  "xlsm",
  "xltx",
  "xltm",
  "ods",
] as const satisfies readonly WorkbookGenerateFormat[];
const PRESENTATION_FORMATS = [
  "pptx",
  "pptm",
  "potx",
  "potm",
  "odp",
] as const satisfies readonly PresentationGenerateFormat[];

function isDocumentFormat(format: OfficeGenerateFormat): format is DocumentGenerateFormat {
  return (DOCUMENT_FORMATS as readonly string[]).includes(format);
}

function isWorkbookFormat(format: OfficeGenerateFormat): format is WorkbookGenerateFormat {
  return (WORKBOOK_FORMATS as readonly string[]).includes(format);
}

function isPresentationFormat(format: OfficeGenerateFormat): format is PresentationGenerateFormat {
  return (PRESENTATION_FORMATS as readonly string[]).includes(format);
}

export async function generateOffice<
  Format extends OfficeGenerateFormat,
  Output extends OutputType = "nodebuffer",
>(
  format: Format,
  options: OfficeOptionsFor[Format],
  output?: Output,
  generationOptions?: OfficeGenerationOptions<Format>,
): Promise<OutputByType[Output]> {
  if (isDocumentFormat(format)) {
    return generateDocument(
      format as DocumentGenerateFormat,
      options as DocumentOptions,
      output,
      generationOptions?.reproducible as ReproducibleGenerationOptions,
    );
  }

  if (isWorkbookFormat(format)) {
    return generateWorkbook(
      format as WorkbookGenerateFormat,
      options as WorkbookOptions,
      output,
      generationOptions?.reproducible as ReproducibleGenerationOptions,
    );
  }

  if (isPresentationFormat(format)) {
    return generatePresentation(
      format as PresentationGenerateFormat,
      options as PresentationOptions,
      output,
      generationOptions?.reproducible as ReproducibleGenerationOptions,
    );
  }

  throw new Error(`${format} is not an office generate format`);
}
