import { readFile, writeFile } from "node:fs/promises";

import type { OutputType, PackerOptions } from "@office-open/core";
export { type OutputType } from "@office-open/core";

import { Document, Packer as DocxPacker } from "@office-open/docx";
import type { PropertiesOptions } from "@office-open/docx";
import { generate as pptxGenerate } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";
import { generate as xlsxGenerate } from "@office-open/xlsx";
import type { WorkbookOptions } from "@office-open/xlsx";

export type GenerateType = "docx" | "pptx" | "xlsx";

export interface GenerateOptions {
  readonly type: GenerateType;
  readonly options: Record<string, unknown>;
  readonly outputType?: OutputType;
}

export async function generate(options: GenerateOptions): Promise<unknown> {
  const { type, options: docOptions, outputType = "nodebuffer" as OutputType } = options;
  const packerOpts = { type: outputType } as PackerOptions<OutputType>;

  switch (type) {
    case "docx": {
      const file = new Document(docOptions as unknown as PropertiesOptions);
      return DocxPacker.pack(file, packerOpts as PackerOptions<"nodebuffer">);
    }
    case "pptx":
      return pptxGenerate(
        docOptions as PresentationOptions,
        packerOpts as PackerOptions<"nodebuffer">,
      );
    case "xlsx":
      return xlsxGenerate(docOptions as WorkbookOptions, packerOpts as PackerOptions<"nodebuffer">);
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
