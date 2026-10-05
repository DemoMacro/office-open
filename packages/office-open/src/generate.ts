import type { OutputByType, OutputType, ReproducibleGenerationOptions } from "@office-open/core";
import type { DocumentOptions } from "@office-open/docx";
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
  odt: DocumentOptions;
  rtf: DocumentOptions;
  xlsx: WorkbookOptions;
  xlsm: WorkbookOptions;
  xltx: WorkbookOptions;
  xltm: WorkbookOptions;
  ods: WorkbookOptions;
  pptx: PresentationOptions;
  pptm: PresentationOptions;
  potx: PresentationOptions;
  potm: PresentationOptions;
  odp: PresentationOptions;
}

type OdfGenerateFormat = "odt" | "ods" | "odp";

/** ODF writers are deterministic today but do not support OPC reproducibility controls. */
export type OfficeGenerationOptions<Format extends OfficeGenerateFormat> = {
  reproducible?: Format extends OdfGenerateFormat ? never : ReproducibleGenerationOptions;
};

export async function generateOffice<
  Format extends OfficeGenerateFormat,
  Output extends OutputType = "nodebuffer",
>(
  format: Format,
  options: OfficeOptionsFor[Format],
  output?: Output,
  generationOptions?: OfficeGenerationOptions<Format>,
): Promise<OutputByType[Output]> {
  const familyGenerators = {
    docx: generateDocument,
    docm: generateDocument,
    dotx: generateDocument,
    dotm: generateDocument,
    odt: generateDocument,
    rtf: generateDocument,
  } as const;

  if (format in familyGenerators) {
    return generateDocument(
      format as DocumentGenerateFormat,
      options as DocumentOptions,
      output,
      generationOptions?.reproducible as DocumentGenerateFormat extends "odt"
        ? never
        : ReproducibleGenerationOptions,
    );
  }

  const workbookFormats = { xlsx: true, xlsm: true, xltx: true, xltm: true, ods: true } as const;
  if (format in workbookFormats) {
    return generateWorkbook(
      format as WorkbookGenerateFormat,
      options as WorkbookOptions,
      output,
      generationOptions?.reproducible as WorkbookGenerateFormat extends "ods"
        ? never
        : ReproducibleGenerationOptions,
    );
  }

  return generatePresentation(
    format as PresentationGenerateFormat,
    options as PresentationOptions,
    output,
    generationOptions?.reproducible as PresentationGenerateFormat extends "odp"
      ? never
      : ReproducibleGenerationOptions,
  );
}
