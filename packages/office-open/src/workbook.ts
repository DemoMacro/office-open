import {
  convertOutput,
  OOXML_PACKAGE_FORMATS,
  type OutputByType,
  type OutputType,
  type PackerOptions,
  type ReproducibleGenerationOptions,
} from "@office-open/core";
import { generateOds, parseOds } from "@office-open/ods";
import { parseWorkbook as parseLegacyWorkbook } from "@office-open/xls";
import type { WorkbookOptions } from "@office-open/xlsx";
import { generateWorkbook, parseWorkbook as parseXlsxWorkbook } from "@office-open/xlsx";

import { detectOfficeFormat } from "./formats";

/** Password accepted by the legacy XLS parser. */
export interface WorkbookFileParseOptions {
  password?: string;
}

export type WorkbookFileParseFormat = "ods" | "xls" | "xlsx" | "xlsm" | "xltx" | "xltm";

export type WorkbookFileGenerateFormat = "xlsx" | "xlsm" | "xltx" | "xltm" | "ods";

const ODS_MIME_TYPE = "application/vnd.oasis.opendocument.spreadsheet";

export async function parseWorkbookFile(
  input: Uint8Array | string,
  options?: WorkbookFileParseOptions,
): Promise<WorkbookOptions> {
  const info = detectOfficeFormat(input);
  const data = typeof input === "string" ? new TextEncoder().encode(input) : input;

  switch (info.format) {
    case "ods":
      return parseOds(data);
    case "xls":
      return parseLegacyWorkbook(data, options);
    case "xlsx":
    case "xlsm":
    case "xltx":
    case "xltm":
      return parseXlsxWorkbook(data);
    default:
      throw new Error(`${info.format} is not a workbook format`);
  }
}

export async function generateWorkbookFile<T extends OutputType = "nodebuffer">(
  format: WorkbookFileGenerateFormat,
  options: WorkbookOptions,
  output?: T,
  reproducible?: ReproducibleGenerationOptions,
): Promise<OutputByType[T]> {
  const outputType = (output ?? "nodebuffer") as T;

  if (format === "ods") {
    return convertOutput(generateOds(options), outputType, ODS_MIME_TYPE);
  }

  return generateWorkbook(options, {
    type: outputType,
    packageVariant: OOXML_PACKAGE_FORMATS[format].variant,
    reproducible,
  } as PackerOptions<T>);
}
