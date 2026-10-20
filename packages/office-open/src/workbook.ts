import {
  toUint8ArrayAsync,
  type DataType,
  type OutputByType,
  type OutputType,
  type PackerOptions,
  type ReproducibleGenerationOptions,
} from "@office-open/core";
import { OOXML_PACKAGE_FORMATS } from "@office-open/core/opc";
import {
  generateWorkbook as generateOdsWorkbook,
  parseWorkbook as parseOdsWorkbook,
  type OdsWorkbookOptions,
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

export type WorkbookPatchRequest<T extends OutputType = "uint8array"> = Omit<
  PatchWorkbookOptions<T>,
  "data" | "outputType"
> & {
  format: WorkbookPatchFormat;
  data: PatchWorkbookOptions<T>["data"];
  outputType?: T;
};

export async function parseWorkbook(
  input: DataType,
  options?: WorkbookParseOptions,
): Promise<WorkbookOptions | OdsWorkbookOptions> {
  const data = await toUint8ArrayAsync(input);
  const info = detectOffice(data);

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
  T extends OutputType = "uint8array",
>(
  format: F,
  options: WorkbookOptions | OdsWorkbookOptions,
  output?: T,
  reproducible?: F extends "ods" ? never : ReproducibleGenerationOptions,
): Promise<OutputByType[T]> {
  const outputType = (output ?? "uint8array") as T;

  if (format === "ods") {
    if (reproducible) throw new Error("Reproducible generation is not supported for ODS");
    return generateOdsWorkbook(options, outputType);
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

export async function patchWorkbook<T extends OutputType = "uint8array">(
  request: WorkbookPatchRequest<T>,
): Promise<OutputByType[T]> {
  const { format, data, outputType = "uint8array" as T, ...patch } = request;
  if (!WORKBOOK_PATCH_FORMATS.includes(format)) {
    throw new Error(`${format} does not support workbook patching`);
  }
  return patchXlsxWorkbook({
    ...(patch as PatchWorkbookOptions<T>),
    data,
    outputType,
  } as PatchWorkbookOptions<T>);
}
