import { parseDocument as parseLegacyDocument } from "@office-open/doc";
import type { DocumentOptions } from "@office-open/docx";
import { parseDocument as parseDocxDocument } from "@office-open/docx";
import { parseOdp, parseOds, parseOdt } from "@office-open/odf";
import { parsePresentation as parseLegacyPresentation } from "@office-open/ppt";
import { parsePresentation as parsePptxPresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";
import { parseRtf } from "@office-open/rtf";
import { parseWorkbook as parseXlsWorkbook } from "@office-open/xls";
import { parseWorkbook as parseXlsxWorkbook } from "@office-open/xlsx";
import type { WorkbookOptions } from "@office-open/xlsx";

import { detectOfficeFormat, type OfficeFormatInfo } from "./formats";

export type ParsedOfficeDocument =
  | { type: OfficeFormatInfo["format"]; options: DocumentOptions }
  | { type: OfficeFormatInfo["format"]; options: WorkbookOptions }
  | { type: OfficeFormatInfo["format"]; options: PresentationOptions };

export async function parseOfficeDocument(
  input: Uint8Array | string,
): Promise<ParsedOfficeDocument> {
  const info = detectOfficeFormat(input);
  const data = typeof input === "string" ? new TextEncoder().encode(input) : input;
  switch (info.format) {
    case "doc":
      return { type: info.format, options: parseLegacyDocument(data) };
    case "docx":
    case "docm":
    case "dotx":
    case "dotm":
      return { type: info.format, options: await parseDocxDocument(data) };
    case "rtf":
      return {
        type: info.format,
        options: parseRtf(typeof input === "string" ? input : new TextDecoder().decode(input)),
      };
    case "odt":
      return { type: info.format, options: parseOdt(data) };
    case "ods":
      return { type: info.format, options: parseOds(data) };
    case "odp":
      return { type: info.format, options: parseOdp(data) };
    case "xls":
      return { type: info.format, options: parseXlsWorkbook(data) };
    case "xlsx":
    case "xlsm":
    case "xltx":
    case "xltm":
      return { type: info.format, options: await parseXlsxWorkbook(data) };
    case "ppt":
      return { type: info.format, options: parseLegacyPresentation(data) };
    case "pptx":
    case "pptm":
    case "potx":
    case "potm":
      return { type: info.format, options: await parsePptxPresentation(data) };
    case "encrypted-ooxml":
      throw new Error("Encrypted OOXML documents are not supported by the unified parser");
  }
}
