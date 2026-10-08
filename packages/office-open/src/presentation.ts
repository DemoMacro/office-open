import {
  convertOutput,
  OOXML_PACKAGE_FORMATS,
  type OutputByType,
  type OutputType,
  type PackerOptions,
  type ReproducibleGenerationOptions,
} from "@office-open/core";
import {
  generatePresentation as generateOdpPresentation,
  parsePresentation as parseOdpPresentation,
} from "@office-open/odp";
import { parsePresentation as parsePptPresentation } from "@office-open/ppt";
import {
  generatePresentation as generatePptxPresentation,
  parsePresentation as parsePptxPresentation,
  patchPresentation as patchPptxPresentation,
  type PatchPresentationOptions,
  type PresentationOptions,
} from "@office-open/pptx";

import { detectOffice, withRequestedPackageVariant } from "./formats";

/** Password accepted by the legacy PPT parser. */
export interface PresentationParseOptions {
  password?: string;
}

export type PresentationParseFormat = "odp" | "ppt" | "pptx" | "pptm" | "potx" | "potm";

export type PresentationGenerateFormat = "pptx" | "pptm" | "potx" | "potm" | "odp";

export type PresentationPatchFormat = "pptx" | "pptm" | "potx" | "potm";

const PRESENTATION_PATCH_FORMATS: readonly PresentationPatchFormat[] = [
  "pptx",
  "pptm",
  "potx",
  "potm",
];

export type PresentationPatchRequest<T extends OutputType = OutputType> = Omit<
  PatchPresentationOptions<T>,
  "data" | "outputType"
> & {
  format: PresentationPatchFormat;
  data: PatchPresentationOptions<T>["data"];
  outputType?: T;
};

const ODP_MIME_TYPE = "application/vnd.oasis.opendocument.presentation";

export async function parsePresentation(
  input: Uint8Array | string,
  options?: PresentationParseOptions,
): Promise<PresentationOptions> {
  const info = detectOffice(input);
  const data = typeof input === "string" ? new TextEncoder().encode(input) : input;

  switch (info.format) {
    case "odp":
      return parseOdpPresentation(data);
    case "ppt":
      return parsePptPresentation(data, options);
    case "pptx":
    case "pptm":
    case "potx":
    case "potm":
      return parsePptxPresentation(data);
    default:
      throw new Error(`${info.format} is not a presentation format`);
  }
}

export async function generatePresentation<
  F extends PresentationGenerateFormat,
  T extends OutputType = "nodebuffer",
>(
  format: F,
  options: PresentationOptions,
  output?: T,
  reproducible?: F extends "odp" ? never : ReproducibleGenerationOptions,
): Promise<OutputByType[T]> {
  const outputType = (output ?? "nodebuffer") as T;

  if (format === "odp") {
    if (reproducible) throw new Error("Reproducible generation is not supported for ODP");
    return convertOutput(generateOdpPresentation(options), outputType, ODP_MIME_TYPE);
  }

  return generatePptxPresentation(
    withRequestedPackageVariant(format as keyof typeof OOXML_PACKAGE_FORMATS, options),
    {
      type: outputType,
      packageVariant: OOXML_PACKAGE_FORMATS[format as keyof typeof OOXML_PACKAGE_FORMATS].variant,
      reproducible,
    } as PackerOptions<T>,
  );
}

export async function patchPresentation<T extends OutputType = OutputType>(
  request: PresentationPatchRequest<T>,
): Promise<OutputByType[T]> {
  const { format, data, outputType = "nodebuffer" as T, ...patch } = request;
  if (!PRESENTATION_PATCH_FORMATS.includes(format)) {
    throw new Error(`${format} does not support presentation patching`);
  }
  return patchPptxPresentation({
    ...(patch as PatchPresentationOptions<T>),
    data,
    outputType,
  } as PatchPresentationOptions<T>);
}
