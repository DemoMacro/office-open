import {
  convertOutput,
  OOXML_PACKAGE_FORMATS,
  type OutputByType,
  type OutputType,
  type PackerOptions,
  type ReproducibleGenerationOptions,
} from "@office-open/core";
import {
  generateWorkbook as generateOdsWorkbook,
  parseWorkbook as parseOdsWorkbook,
} from "@office-open/ods";
import { parseWorkbook as parseXlsWorkbook } from "@office-open/xls";
import {
  generateWorkbook as generateXlsxWorkbook,
  parseWorkbook as parseXlsxWorkbook,
  patchWorkbook as patchXlsxWorkbook,
  type PatchWorkbookOptions,
  type WorkbookOptions,
} from "@office-open/xlsx";

import { detectOffice, withRequestedPackageVariant } from "./formats";

/** Password accepted by the legacy XLS parser. */
export interface WorkbookParseOptions {
  password?: string;
}

export type WorkbookParseFormat = "ods" | "xls" | "xlsx" | "xlsm" | "xltx" | "xltm";

export type WorkbookGenerateFormat = "xlsx" | "xlsm" | "xltx" | "xltm" | "ods";

export type WorkbookPatchFormat = "xlsx" | "xlsm" | "xltx" | "xltm";

const WORKBOOK_PATCH_FORMATS: readonly WorkbookPatchFormat[] = ["xlsx", "xlsm", "xltx", "xltm"];

export type WorkbookPatchRequest<T extends OutputType = OutputType> = Omit<
  PatchWorkbookOptions<T>,
  "data" | "outputType"
> & {
  format: WorkbookPatchFormat;
  data: PatchWorkbookOptions<T>["data"];
  outputType?: T;
};

const ODS_MIME_TYPE = "application/vnd.oasis.opendocument.spreadsheet";

export async function parseWorkbook(
  input: Uint8Array | string,
  options?: WorkbookParseOptions,
): Promise<WorkbookOptions> {
  const info = detectOffice(input);
  const data = typeof input === "string" ? new TextEncoder().encode(input) : input;

  switch (info.format) {
    case "ods":
      return parseOdsWorkbook(data);
    case "xls":
      return parseXlsWorkbook(data, options);
    case "xlsx":
    case "xlsm":
    case "xltx":
    case "xltm":
      return parseXlsxWorkbook(data);
    default:
      throw new Error(`${info.format} is not a workbook format`);
  }
}

export async function generateWorkbook<
  F extends WorkbookGenerateFormat,
  T extends OutputType = "nodebuffer",
>(
  format: F,
  options: WorkbookOptions,
  output?: T,
  reproducible?: F extends "ods" ? never : ReproducibleGenerationOptions,
): Promise<OutputByType[T]> {
  const outputType = (output ?? "nodebuffer") as T;

  if (format === "ods") {
    if (reproducible) throw new Error("Reproducible generation is not supported for ODS");
    return convertOutput(generateOdsWorkbook(options), outputType, ODS_MIME_TYPE);
  }

  return generateXlsxWorkbook(
    withRequestedPackageVariant(format as keyof typeof OOXML_PACKAGE_FORMATS, options),
    {
      type: outputType,
      packageVariant: OOXML_PACKAGE_FORMATS[format as keyof typeof OOXML_PACKAGE_FORMATS].variant,
      reproducible,
    } as PackerOptions<T>,
  );
}

export async function patchWorkbook<T extends OutputType = OutputType>(
  request: WorkbookPatchRequest<T>,
): Promise<OutputByType[T]> {
  const { format, data, outputType = "nodebuffer" as T, ...patch } = request;
  if (!WORKBOOK_PATCH_FORMATS.includes(format)) {
    throw new Error(`${format} does not support workbook patching`);
  }
  return patchXlsxWorkbook({
    ...(patch as PatchWorkbookOptions<T>),
    data,
    outputType,
  } as PatchWorkbookOptions<T>);
}
