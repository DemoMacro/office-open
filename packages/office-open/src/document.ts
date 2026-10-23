import {
  convertOutput,
  unsupportedPatchFormat,
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
  type DocumentPatchOptions,
} from "@office-open/docx";
import {
  patchDocument as patchOdtDocument,
  generateDocument as generateOdtDocument,
  parseDocument as parseOdtDocument,
  type OdtPatchOptions,
  type OdtDocumentOptions,
} from "@office-open/odt";
import {
  generateDocument as generateRtfDocument,
  parseDocument as parseRtfDocument,
  patchDocument as patchRtfDocument,
  type RtfPatchOptions,
} from "@office-open/rtf";

import { detectOffice, withRequestedPackageVariant } from "./formats";

/** Password accepted by the legacy DOC parser. */
export interface DocumentParseOptions {
  password?: string;
}

export type DocumentParseFormat = "doc" | "docx" | "docm" | "dotx" | "dotm" | "rtf" | "odt";

export type DocumentGenerateFormat = "docx" | "docm" | "dotx" | "dotm" | "odt" | "rtf";

type DocumentOoxmlPatchFormat = "docx" | "docm" | "dotx" | "dotm";

export type DocumentPatchFormat = DocumentOoxmlPatchFormat | "rtf" | "odt";

const DOCUMENT_PATCH_FORMATS: readonly DocumentPatchFormat[] = [
  "docx",
  "docm",
  "dotx",
  "dotm",
  "rtf",
  "odt",
];
type DocumentOoxmlPatchRequest<T extends OutputType = "uint8array"> = Omit<
  DocumentPatchOptions<T>,
  "outputType"
> & {
  format: DocumentOoxmlPatchFormat;
  outputType?: T;
};

type DocumentRtfPatchRequest<T extends OutputType = "uint8array"> = Omit<
  RtfPatchOptions<T>,
  "outputType"
> & {
  format: "rtf";
  outputType?: T;
};

type DocumentOdtPatchRequest<T extends OutputType = "uint8array"> = Omit<
  OdtPatchOptions<T>,
  "outputType"
> & {
  format: "odt";
  outputType?: T;
};

export type DocumentPatchRequest<T extends OutputType = "uint8array"> =
  | DocumentOoxmlPatchRequest<T>
  | DocumentRtfPatchRequest<T>
  | DocumentOdtPatchRequest<T>;

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
      return parseRtfDocument(data) as DocumentOptions;
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
  const { format } = request;
  if (!DOCUMENT_PATCH_FORMATS.includes(format)) {
    throw unsupportedPatchFormat(format, "document");
  }
  if (format === "docx" || format === "docm" || format === "dotx" || format === "dotm") {
    const {
      data,
      outputType = "uint8array" as T,
      packageVariant,
      placeholderDelimiters,
      commands,
    } = request;
    return patchDocxDocument({ commands, data, outputType, packageVariant, placeholderDelimiters });
  }
  if (format === "rtf") {
    const { data, outputType = "uint8array" as T, placeholderDelimiters, commands } = request;
    return patchRtfDocument({ commands, data, outputType, placeholderDelimiters });
  }
  if (format === "odt") {
    const { data, outputType = "uint8array" as T, placeholderDelimiters, commands } = request;
    return patchOdtDocument({ commands, data, outputType, placeholderDelimiters });
  }
  throw unsupportedPatchFormat(format, "document");
}
