import {
  convertOutput,
  OOXML_PACKAGE_FORMATS,
  type OutputByType,
  type OutputType,
  type PackerOptions,
  type ReproducibleGenerationOptions,
} from "@office-open/core";
import { parseDocument as parseLegacyDocument } from "@office-open/doc";
import type { DocumentOptions } from "@office-open/docx";
import { generateDocument, parseDocument as parseDocxDocument } from "@office-open/docx";
import { generateOdt, parseOdt } from "@office-open/odt";
import { parseRtf } from "@office-open/rtf";

import { detectOfficeFormat } from "./formats";

/** Password accepted by the legacy DOC parser. */
export interface DocumentFileParseOptions {
  password?: string;
}

export type DocumentFileParseFormat = "doc" | "docx" | "docm" | "dotx" | "dotm" | "rtf" | "odt";

export type DocumentFileGenerateFormat = "docx" | "docm" | "dotx" | "dotm" | "odt";

const ODT_MIME_TYPE = "application/vnd.oasis.opendocument.text";

export async function parseDocumentFile(
  input: Uint8Array | string,
  options?: DocumentFileParseOptions,
): Promise<DocumentOptions> {
  const info = detectOfficeFormat(input);
  const data = typeof input === "string" ? new TextEncoder().encode(input) : input;

  switch (info.format) {
    case "doc":
      return parseLegacyDocument(data, options);
    case "docx":
    case "docm":
    case "dotx":
    case "dotm":
      return parseDocxDocument(data);
    case "rtf":
      return parseRtf(typeof input === "string" ? input : new TextDecoder().decode(input));
    case "odt":
      return parseOdt(data);
    default:
      throw new Error(`${info.format} is not a document format`);
  }
}

export async function generateDocumentFile<
  F extends DocumentFileGenerateFormat,
  T extends OutputType = "nodebuffer",
>(
  format: F,
  options: DocumentOptions,
  output?: T,
  reproducible?: F extends "odt" ? never : ReproducibleGenerationOptions,
): Promise<OutputByType[T]> {
  const outputType = (output ?? "nodebuffer") as T;

  if (format === "odt") {
    if (reproducible) throw new Error("Reproducible generation is not supported for ODT");
    return convertOutput(generateOdt(options), outputType, ODT_MIME_TYPE);
  }

  return generateDocument(options, {
    type: outputType,
    packageVariant: OOXML_PACKAGE_FORMATS[format as keyof typeof OOXML_PACKAGE_FORMATS].variant,
    reproducible,
  } as PackerOptions<T>);
}
