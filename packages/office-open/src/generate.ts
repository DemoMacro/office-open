import { readFile, writeFile } from "node:fs/promises";

import type {
  OoxmlPackageFormat,
  OutputType,
  PackerOptions,
  ReproducibleGenerationOptions,
} from "@office-open/core";
import { OOXML_PACKAGE_FORMATS } from "@office-open/core";
export { type OutputType } from "@office-open/core";

import { generateDocument } from "@office-open/docx";
import type { DocumentOptions } from "@office-open/docx";
import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";
import { generateWorkbook } from "@office-open/xlsx";
import type { WorkbookOptions } from "@office-open/xlsx";

export type GenerateType = OoxmlPackageFormat;

/** Map from type string to the corresponding options type. */
export interface GenerateOptionsMap {
  docx: DocumentOptions;
  docm: DocumentOptions;
  dotx: DocumentOptions;
  dotm: DocumentOptions;
  pptx: PresentationOptions;
  pptm: PresentationOptions;
  potx: PresentationOptions;
  potm: PresentationOptions;
  xlsx: WorkbookOptions;
  xlsm: WorkbookOptions;
  xltx: WorkbookOptions;
  xltm: WorkbookOptions;
}

export interface GenerateOptions<T extends GenerateType = GenerateType> {
  type: T;
  options: GenerateOptionsMap[T];
  outputType?: OutputType;
  /** Opt-in reproducible generation (see PackerOptions.reproducible). */
  reproducible?: ReproducibleGenerationOptions;
}

export async function generate<T extends GenerateType>(
  options: GenerateOptions<T>,
): Promise<unknown> {
  const { type, options: docOptions, outputType = "nodebuffer" as OutputType } = options;
  const packageFormat = OOXML_PACKAGE_FORMATS[type];
  const packerOpts = {
    type: outputType,
    packageVariant: packageFormat.variant,
    reproducible: options.reproducible,
  } as PackerOptions<OutputType>;

  switch (packageFormat.family) {
    case "wordprocessing":
      return generateDocument(
        docOptions as DocumentOptions,
        packerOpts as PackerOptions<"nodebuffer">,
      );
    case "presentation":
      return generatePresentation(
        docOptions as PresentationOptions,
        packerOpts as PackerOptions<"nodebuffer">,
      );
    case "spreadsheet":
      return generateWorkbook(
        docOptions as WorkbookOptions,
        packerOpts as PackerOptions<"nodebuffer">,
      );
  }
}

export async function parseInput(input: string): Promise<Record<string, unknown>> {
  const trimmed = input.trim();
  if (trimmed.startsWith("{") || trimmed.startsWith("[")) {
    return JSON.parse(trimmed) as Record<string, unknown>;
  }
  const content = await readFile(trimmed, "utf-8");
  return JSON.parse(content) as Record<string, unknown>;
}

export async function generateToFile(outputPath: string, options: GenerateOptions): Promise<void> {
  const buffer = (await generate({ ...options, outputType: "nodebuffer" })) as Buffer;
  await writeFile(outputPath, buffer);
}
