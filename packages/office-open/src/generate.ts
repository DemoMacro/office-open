import { readFile, writeFile } from "node:fs/promises";

import type { OutputType } from "@office-open/core";
export { type OutputType } from "@office-open/core";

import { Document, Packer as DocxPacker } from "@office-open/docx";
import type { PropertiesOptions } from "@office-open/docx";
import { Presentation, Packer as PptxPacker } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";
import { Workbook, Packer as XlsxPacker } from "@office-open/xlsx";
import type { WorkbookOptions } from "@office-open/xlsx";

export type GenerateType = "docx" | "pptx" | "xlsx";

export interface GenerateOptions {
  readonly type: GenerateType;
  readonly options: Record<string, unknown>;
  readonly outputType?: OutputType;
}

const PACKERS = {
  docx: {
    newFile: (opts: unknown) => new Document(opts as unknown as PropertiesOptions),
    pack: (file: unknown, type: OutputType) => DocxPacker.pack(file as Document, type),
  },
  pptx: {
    newFile: (opts: unknown) => new Presentation(opts as unknown as PresentationOptions),
    pack: (file: unknown, type: OutputType) => PptxPacker.pack(file as Presentation, type),
  },
  xlsx: {
    newFile: (opts: unknown) => new Workbook(opts as unknown as WorkbookOptions),
    pack: (file: unknown, type: OutputType) => XlsxPacker.pack(file as Workbook, type),
  },
} as const satisfies Record<
  GenerateType,
  {
    newFile: (opts: unknown) => unknown;
    pack: (file: unknown, type: OutputType) => Promise<unknown>;
  }
>;

export async function generate(options: GenerateOptions): Promise<unknown> {
  const { type, options: docOptions, outputType = "nodebuffer" as OutputType } = options;
  const { newFile, pack } = PACKERS[type];
  const file = newFile(docOptions);
  return pack(file, outputType);
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
