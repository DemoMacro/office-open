import {
  convertOutput,
  OOXML_PACKAGE_FORMATS,
  type OutputByType,
  type OutputType,
  type PackerOptions,
  type ReproducibleGenerationOptions,
} from "@office-open/core";
import { generateOdp, parseOdp } from "@office-open/odp";
import { parsePresentation as parseLegacyPresentation } from "@office-open/ppt";
import type { PresentationOptions } from "@office-open/pptx";
import {
  generatePresentation,
  parsePresentation as parsePptxPresentation,
} from "@office-open/pptx";

import { detectOfficeFormat } from "./formats";

/** Password accepted by the legacy PPT parser. */
export interface PresentationFileParseOptions {
  password?: string;
}

export type PresentationFileParseFormat = "odp" | "ppt" | "pptx" | "pptm" | "potx" | "potm";

export type PresentationFileGenerateFormat = "pptx" | "pptm" | "potx" | "potm" | "odp";

const ODP_MIME_TYPE = "application/vnd.oasis.opendocument.presentation";

export async function parsePresentationFile(
  input: Uint8Array | string,
  options?: PresentationFileParseOptions,
): Promise<PresentationOptions> {
  const info = detectOfficeFormat(input);
  const data = typeof input === "string" ? new TextEncoder().encode(input) : input;

  switch (info.format) {
    case "odp":
      return parseOdp(data);
    case "ppt":
      return parseLegacyPresentation(data, options);
    case "pptx":
    case "pptm":
    case "potx":
    case "potm":
      return parsePptxPresentation(data);
    default:
      throw new Error(`${info.format} is not a presentation format`);
  }
}

export async function generatePresentationFile<
  F extends PresentationFileGenerateFormat,
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
    return convertOutput(generateOdp(options), outputType, ODP_MIME_TYPE);
  }

  return generatePresentation(options, {
    type: outputType,
    packageVariant: OOXML_PACKAGE_FORMATS[format as keyof typeof OOXML_PACKAGE_FORMATS].variant,
    reproducible,
  } as PackerOptions<T>);
}
