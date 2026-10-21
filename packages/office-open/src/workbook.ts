import {
  toUint8ArrayAsync,
  unsupportedPatchFormat,
  type DataType,
  type OutputByType,
  type OutputType,
  type PackerOptions,
  type ReproducibleGenerationOptions,
} from "@office-open/core";
import { OOXML_PACKAGE_FORMATS } from "@office-open/core/opc";
import {
  patchWorkbook as patchOdsWorkbook,
  generateWorkbook as generateOdsWorkbook,
  parseWorkbook as parseOdsWorkbook,
  type OdsPatchOptions,
  type OdsWorkbookOptions,
} from "@office-open/ods";
import { parseWorkbook as parseXlsWorkbook } from "@office-open/xls";
import {
  generateWorkbook as generateXlsxWorkbook,
  parseWorkbook as parseXlsxWorkbook,
  patchWorkbook as patchXlsxWorkbook,
  type WorkbookPatchOptions,
  type WorkbookOptions,
} from "@office-open/xlsx";

import { detectOffice, withRequestedPackageVariant } from "./formats";

/** Password accepted by the legacy XLS parser. */
export interface WorkbookParseOptions {
  password?: string;
}

export type WorkbookParseFormat = "ods" | "xls" | "xlsx" | "xlsm" | "xltx" | "xltm";

export type WorkbookGenerateFormat = "xlsx" | "xlsm" | "xltx" | "xltm" | "ods";

type WorkbookOoxmlPatchFormat = "xlsx" | "xlsm" | "xltx" | "xltm";

export type WorkbookPatchFormat = WorkbookOoxmlPatchFormat | "ods";

const WORKBOOK_PATCH_FORMATS: readonly WorkbookPatchFormat[] = [
  "xlsx",
  "xlsm",
  "xltx",
  "xltm",
  "ods",
];
type WorkbookOoxmlPatchRequest<T extends OutputType = "uint8array"> = Omit<
  WorkbookPatchOptions<T>,
  "outputType"
> & {
  format: WorkbookOoxmlPatchFormat;
  outputType?: T;
};

type WorkbookOdsPatchRequest<T extends OutputType = "uint8array"> = Omit<
  OdsPatchOptions<T>,
  "outputType"
> & {
  format: "ods";
  outputType?: T;
};

export type WorkbookPatchRequest<T extends OutputType = "uint8array"> =
  | WorkbookOoxmlPatchRequest<T>
  | WorkbookOdsPatchRequest<T>;

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
  const { format } = request;
  if (!WORKBOOK_PATCH_FORMATS.includes(format)) {
    throw unsupportedPatchFormat(format, "workbook");
  }
  if (format === "xlsx" || format === "xlsm" || format === "xltx" || format === "xltm") {
    const {
      data,
      outputType = "uint8array" as T,
      packageVariant,
      placeholderDelimiters,
      commands,
    } = request;
    return patchXlsxWorkbook({
      commands,
      data,
      outputType,
      packageVariant,
      placeholderDelimiters,
    });
  }
  if (format === "ods") {
    const { data, outputType = "uint8array" as T, placeholderDelimiters, commands } = request;
    return patchOdsWorkbook({ commands, data, outputType, placeholderDelimiters });
  }
  throw unsupportedPatchFormat(format, "workbook");
}
