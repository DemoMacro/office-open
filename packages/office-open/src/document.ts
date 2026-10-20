import {
  convertOutput,
  toUint8ArrayAsync,
  type DataType,
  type OutputByType,
  type OutputType,
  type PackerOptions,
  type ReproducibleGenerationOptions,
} from "@office-open/core";
import { OOXML_PACKAGE_FORMATS } from "@office-open/core/opc";
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

export type DocumentPatchRequest<T extends OutputType = "uint8array"> = Omit<
  PatchDocumentOptions<T>,
  "data" | "outputType"
> & {
  format: DocumentPatchFormat;
  data: PatchDocumentOptions<T>["data"];
  outputType?: T;
};

const RTF_MIME_TYPE = "application/rtf";

export async function parseDocument(
  input: DataType,
  options?: DocumentParseOptions,
): Promise<DocumentOptions | OdtDocumentOptions> {
  const data = await toUint8ArrayAsync(input);
  const info = detectOffice(data);

  switch (info.format) {
    case "doc":
      return parseDocDocument(data, options) as DocumentOptions;
    case "docx":
    case "docm":
    case "dotx":
    case "dotm":
      return parseDocxDocument(data);
    case "rtf":
      return parseRtfDocument(new TextDecoder().decode(data)) as DocumentOptions;
    case "odt":
      return parseOdtDocument(data);
    default:
      throw new Error(`${info.format} is not a document format`);
  }
}

export async function generateDocument<
  F extends DocumentGenerateFormat,
  T extends OutputType = "uint8array",
>(
  format: F,
  options: DocumentOptions | OdtDocumentOptions,
  output?: T,
  reproducible?: F extends "odt" ? never : ReproducibleGenerationOptions,
): Promise<OutputByType[T]> {
  const outputType = (output ?? "uint8array") as T;

  if (format === "odt") {
    if (reproducible) throw new Error("Reproducible generation is not supported for ODT");
    return generateOdtDocument(options as OdtDocumentOptions, outputType);
  }

  if (format === "rtf") {
    return convertOutput(
      generateRtfDocument(options as Parameters<typeof generateRtfDocument>[0]),
      outputType,
      RTF_MIME_TYPE,
    );
  }

  return generateDocxDocument(
    withRequestedPackageVariant(
      format as keyof typeof OOXML_PACKAGE_FORMATS,
      options as DocumentOptions,
    ),
    {
      type: outputType,
      packageVariant: OOXML_PACKAGE_FORMATS[format as keyof typeof OOXML_PACKAGE_FORMATS].variant,
      reproducible,
    } as PackerOptions<T>,
  );
}

export async function patchDocument<T extends OutputType = "uint8array">(
  request: DocumentPatchRequest<T>,
): Promise<OutputByType[T]> {
  const { format, data, outputType = "uint8array" as T, ...patch } = request;
  if (!DOCUMENT_PATCH_FORMATS.includes(format)) {
    throw new Error(`${format} does not support document patching`);
  }
  return patchDocxDocument({
    ...(patch as PatchDocumentOptions<T>),
    data,
    outputType,
  } as PatchDocumentOptions<T>);
}
