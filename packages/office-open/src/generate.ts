import { readFile, writeFile } from "node:fs/promises";

import type { OutputByType, OutputType, ReproducibleGenerationOptions } from "@office-open/core";
export type { OutputType } from "@office-open/core";

import type { DocumentOptions } from "@office-open/docx";
import type { PresentationOptions } from "@office-open/pptx";
import type { WorkbookOptions } from "@office-open/xlsx";

import { generateDocumentFile, type DocumentFileGenerateFormat } from "./document";
import { generatePresentationFile, type PresentationFileGenerateFormat } from "./presentation";
import { generateWorkbookFile, type WorkbookFileGenerateFormat } from "./workbook";

export type OfficeGenerateFormat =
  | DocumentFileGenerateFormat
  | WorkbookFileGenerateFormat
  | PresentationFileGenerateFormat;

export type GenerateType = OfficeGenerateFormat;

/** Map from type string to the corresponding options type. */
export interface GenerateOptionsMap {
  docx: DocumentOptions;
  docm: DocumentOptions;
  dotx: DocumentOptions;
  dotm: DocumentOptions;
  odt: DocumentOptions;
  pptx: PresentationOptions;
  pptm: PresentationOptions;
  potx: PresentationOptions;
  potm: PresentationOptions;
  odp: PresentationOptions;
  rtf: DocumentOptions;
  xlsx: WorkbookOptions;
  xlsm: WorkbookOptions;
  xltx: WorkbookOptions;
  xltm: WorkbookOptions;
  ods: WorkbookOptions;
}

type OdfGenerateFormat = "odt" | "ods" | "odp";

/** ODF writers are deterministic today but do not support OPC reproducibility controls. */
type FormatReproducibleOptions<F extends GenerateType> = F extends OdfGenerateFormat
  ? never
  : ReproducibleGenerationOptions;

interface OfficeGenerationOptions<F extends GenerateType> {
  reproducible?: FormatReproducibleOptions<F>;
}

export async function generateOfficeDocument<
  F extends GenerateType,
  T extends OutputType = "nodebuffer",
>(
  format: F,
  options: GenerateOptionsMap[F],
  output?: T,
  generationOptions?: OfficeGenerationOptions<F>,
): Promise<OutputByType[T]> {
  if (format === "odt") {
    return generateDocumentFile(
      format,
      options as DocumentOptions,
      output,
      generationOptions?.reproducible,
    );
  }

  if (format === "rtf") {
    return generateDocumentFile(format, options as DocumentOptions, output);
  }
  if (format === "ods") {
    return generateWorkbookFile(
      format,
      options as WorkbookOptions,
      output,
      generationOptions?.reproducible,
    );
  }
  if (format === "odp") {
    return generatePresentationFile(
      format,
      options as PresentationOptions,
      output,
      generationOptions?.reproducible,
    );
  }

  const families = {
    docx: "document",
    docm: "document",
    dotx: "document",
    dotm: "document",
    pptx: "presentation",
    pptm: "presentation",
    potx: "presentation",
    potm: "presentation",
    xlsx: "workbook",
    xlsm: "workbook",
    xltx: "workbook",
    xltm: "workbook",
  } as const;

  if (families[format as keyof typeof families] === "document") {
    return generateDocumentFile(
      format as DocumentFileGenerateFormat,
      options as DocumentOptions,
      output,
      generationOptions?.reproducible,
    );
  }
  if (families[format as keyof typeof families] === "workbook") {
    return generateWorkbookFile(
      format as WorkbookFileGenerateFormat,
      options as WorkbookOptions,
      output,
      generationOptions?.reproducible,
    );
  }
  return generatePresentationFile(
    format as PresentationFileGenerateFormat,
    options as PresentationOptions,
    output,
    generationOptions?.reproducible,
  );
}

export interface GenerateOptions<T extends GenerateType = GenerateType> {
  type: T;
  options: GenerateOptionsMap[T];
  outputType?: OutputType;
  /** Opt-in OPC reproducibility controls; ODF formats reject this option. */
  reproducible?: FormatReproducibleOptions<T>;
}

export async function generate<T extends GenerateType>(
  options: GenerateOptions<T>,
): Promise<unknown> {
  const { type, options: docOptions, outputType = "nodebuffer" as OutputType } = options;
  return generateOfficeDocument(type, docOptions, outputType, {
    reproducible: options.reproducible,
  });
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
