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
  patchPresentation as patchOdpPresentation,
  generatePresentation as generateOdpPresentation,
  parsePresentation as parseOdpPresentation,
  type OdpPatchOptions,
  type OdpPresentationOptions,
} from "@office-open/odp";
import { parsePresentation as parsePptPresentation } from "@office-open/ppt";
import {
  generatePresentation as generatePptxPresentation,
  parsePresentation as parsePptxPresentation,
  patchPresentation as patchPptxPresentation,
  type PresentationPatchOptions,
  type PresentationOptions,
} from "@office-open/pptx";

import { detectOffice, withRequestedPackageVariant } from "./formats";

/** Password accepted by the legacy PPT parser. */
export interface PresentationParseOptions {
  password?: string;
}

export type PresentationParseFormat = "odp" | "ppt" | "pptx" | "pptm" | "potx" | "potm";

export type PresentationGenerateFormat = "pptx" | "pptm" | "potx" | "potm" | "odp";

type PresentationOoxmlPatchFormat = "pptx" | "pptm" | "potx" | "potm";

export type PresentationPatchFormat = PresentationOoxmlPatchFormat | "odp";

const PRESENTATION_PATCH_FORMATS: readonly PresentationPatchFormat[] = [
  "pptx",
  "pptm",
  "potx",
  "potm",
  "odp",
];
type PresentationOoxmlPatchRequest<T extends OutputType = "uint8array"> = Omit<
  PresentationPatchOptions<T>,
  "outputType"
> & {
  format: PresentationOoxmlPatchFormat;
  outputType?: T;
};

type PresentationOdpPatchRequest<T extends OutputType = "uint8array"> = Omit<
  OdpPatchOptions<T>,
  "outputType"
> & {
  format: "odp";
  outputType?: T;
};

export type PresentationPatchRequest<T extends OutputType = "uint8array"> =
  | PresentationOoxmlPatchRequest<T>
  | PresentationOdpPatchRequest<T>;

export async function parsePresentation(
  input: DataType,
  options?: PresentationParseOptions,
): Promise<PresentationOptions | OdpPresentationOptions> {
  const data = await toUint8ArrayAsync(input);
  const info = detectOffice(data);

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
  T extends OutputType = "uint8array",
>(
  format: F,
  options: PresentationOptions | OdpPresentationOptions,
  output?: T,
  reproducible?: F extends "odp" ? never : ReproducibleGenerationOptions,
): Promise<OutputByType[T]> {
  const outputType = (output ?? "uint8array") as T;

  if (format === "odp") {
    if (reproducible) throw new Error("Reproducible generation is not supported for ODP");
    return generateOdpPresentation(options, outputType);
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

export async function patchPresentation<T extends OutputType = "uint8array">(
  request: PresentationPatchRequest<T>,
): Promise<OutputByType[T]> {
  const { format } = request;
  if (!PRESENTATION_PATCH_FORMATS.includes(format)) {
    throw unsupportedPatchFormat(format, "presentation");
  }
  if (format === "pptx" || format === "pptm" || format === "potx" || format === "potm") {
    const {
      data,
      outputType = "uint8array" as T,
      packageVariant,
      placeholderDelimiters,
      commands,
    } = request;
    return patchPptxPresentation({
      commands,
      data,
      outputType,
      packageVariant,
      placeholderDelimiters,
    });
  }
  if (format === "odp") {
    const { data, outputType = "uint8array" as T, placeholderDelimiters, commands } = request;
    return patchOdpPresentation({ commands, data, outputType, placeholderDelimiters });
  }
  throw unsupportedPatchFormat(format, "presentation");
}
