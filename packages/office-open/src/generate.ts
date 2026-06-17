import { readFile, writeFile } from "node:fs/promises";

import type { OutputType, PackerOptions } from "@office-open/core";
export { type OutputType } from "@office-open/core";

import { generateDocument } from "@office-open/docx";
import type { DocumentOptions } from "@office-open/docx";
import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";
import { generateWorkbook } from "@office-open/xlsx";
import type { WorkbookOptions } from "@office-open/xlsx";

export type GenerateType = "docx" | "pptx" | "xlsx";

/** Map from type string to the corresponding options type. */
export interface GenerateOptionsMap {
  docx: DocumentOptions;
  pptx: PresentationOptions;
  xlsx: WorkbookOptions;
}

export interface GenerateOptions<T extends GenerateType = GenerateType> {
  type: T;
  options: GenerateOptionsMap[T];
  outputType?: OutputType;
}

export async function generate<T extends GenerateType>(
  options: GenerateOptions<T>,
): Promise<unknown> {
  const { type, options: docOptions, outputType = "nodebuffer" as OutputType } = options;
  const packerOpts = { type: outputType } as PackerOptions<OutputType>;

  switch (type) {
    case "docx":
      return generateDocument(
        docOptions as DocumentOptions,
        packerOpts as PackerOptions<"nodebuffer">,
      );
    case "pptx":
      return generatePresentation(
        docOptions as PresentationOptions,
        packerOpts as PackerOptions<"nodebuffer">,
      );
    case "xlsx":
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
