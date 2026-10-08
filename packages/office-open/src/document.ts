import {
  convertOutput,
  OOXML_PACKAGE_FORMATS,
  type OutputByType,
  type OutputType,
  type PackerOptions,
  type ReproducibleGenerationOptions,
} from "@office-open/core";
import { parseDocument as parseDocDocument } from "@office-open/doc";
import {
  generateDocument as generateDocxDocument,
  parseDocument as parseDocxDocument,
  patchDocument as patchDocxDocument,
  type DocumentOptions,
  type PatchDocumentOptions,
} from "@office-open/docx";
import {
  generateDocument as generateOdtDocument,
  parseDocument as parseOdtDocument,
  type OdtDocumentOptions,
} from "@office-open/odt";
import {
  generateDocument as generateRtfDocument,
  parseDocument as parseRtfDocument,
} from "@office-open/rtf";

import { detectOffice, withRequestedPackageVariant } from "./formats";

/** Password accepted by the legacy DOC parser. */
export interface DocumentParseOptions {
  password?: string;
}

export type DocumentParseFormat = "doc" | "docx" | "docm" | "dotx" | "dotm" | "rtf" | "odt";

export type DocumentGenerateFormat = "docx" | "docm" | "dotx" | "dotm" | "odt" | "rtf";

export type DocumentPatchFormat = "docx" | "docm" | "dotx" | "dotm";

const DOCUMENT_PATCH_FORMATS: readonly DocumentPatchFormat[] = ["docx", "docm", "dotx", "dotm"];

export type DocumentPatchRequest<T extends OutputType = OutputType> = Omit<
  PatchDocumentOptions<T>,
  "data" | "outputType"
> & {
  format: DocumentPatchFormat;
  data: PatchDocumentOptions<T>["data"];
  outputType?: T;
};

const ODT_MIME_TYPE = "application/vnd.oasis.opendocument.text";
const RTF_MIME_TYPE = "application/rtf";

export async function parseDocument(
  input: Uint8Array | string,
  options?: DocumentParseOptions,
): Promise<DocumentOptions> {
  const info = detectOffice(input);
  const data = typeof input === "string" ? new TextEncoder().encode(input) : input;

  switch (info.format) {
    case "doc":
      return parseDocDocument(data, options);
    case "docx":
    case "docm":
    case "dotx":
    case "dotm":
      return parseDocxDocument(data);
    case "rtf":
      return parseRtfDocument(typeof input === "string" ? input : new TextDecoder().decode(input));
    case "odt":
      return parseOdtDocument(data) as DocumentOptions;
    default:
      throw new Error(`${info.format} is not a document format`);
  }
}

export async function generateDocument<
  F extends DocumentGenerateFormat,
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
    return convertOutput(
      generateOdtDocument(options as OdtDocumentOptions),
      outputType,
      ODT_MIME_TYPE,
    );
  }

  if (format === "rtf") {
    return convertOutput(generateRtfDocument(options), outputType, RTF_MIME_TYPE);
  }

  return generateDocxDocument(
    withRequestedPackageVariant(format as keyof typeof OOXML_PACKAGE_FORMATS, options),
    {
      type: outputType,
      packageVariant: OOXML_PACKAGE_FORMATS[format as keyof typeof OOXML_PACKAGE_FORMATS].variant,
      reproducible,
    } as PackerOptions<T>,
  );
}

export async function patchDocument<T extends OutputType = OutputType>(
  request: DocumentPatchRequest<T>,
): Promise<OutputByType[T]> {
  const { format, data, outputType = "nodebuffer" as T, ...patch } = request;
  if (!DOCUMENT_PATCH_FORMATS.includes(format)) {
    throw new Error(`${format} does not support document patching`);
  }
  return patchDocxDocument({
    ...(patch as PatchDocumentOptions<T>),
    data,
    outputType,
  } as PatchDocumentOptions<T>);
}
