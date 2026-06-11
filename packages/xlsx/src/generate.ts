/**
 * Pure function API for generating XLSX files.
 *
 * @module
 */

import { createPacker, OoxmlMimeType } from "@office-open/core";
import type { OutputByType, OutputType, PackerOptions } from "@office-open/core";
import type { WorkbookOptions } from "@parts/file";

import { compileWorkbook } from "./compiler";

/** @internal Packer instance for XLSX generation. */
const Packer = createPacker<WorkbookOptions>({
  compile: (options, overrides, mediaLevel) => compileWorkbook(options, overrides, mediaLevel),
  mimeType: OoxmlMimeType.XLSX,
});

/**
 * Generate an XLSX file from pure JSON options.
 *
 * The output format is controlled by `packerOptions.type` (default: `"nodebuffer"` → Buffer).
 * For synchronous generation, use {@link generateWorkbookSync}. For streaming, use {@link generateWorkbookStream}.
 *
 * @param options - Workbook options (worksheets, styles, etc.)
 * @param packerOptions - Optional packer configuration (type, compression, overrides, etc.)
 *
 * @example
 * ```typescript
 * import { generateWorkbook } from "@office-open/xlsx";
 *
 * const buffer = await generateWorkbook({ worksheets: [...] });
 * const bytes = await generateWorkbook({ worksheets: [...] }, { type: "uint8array" });
 * const blob = await generateWorkbook({ worksheets: [...] }, { type: "blob" });
 * ```
 */
export function generateWorkbook<T extends OutputType = "nodebuffer">(
  options: WorkbookOptions,
  packerOptions?: PackerOptions<T>,
): Promise<OutputByType[T]> {
  return Packer.pack(options, packerOptions) as Promise<OutputByType[T]>;
}

/**
 * Synchronously generate an XLSX file from pure JSON options.
 */
export function generateWorkbookSync<T extends OutputType = "nodebuffer">(
  options: WorkbookOptions,
  packerOptions?: PackerOptions<T>,
): OutputByType[T] {
  return Packer.packSync(options, packerOptions) as OutputByType[T];
}

/**
 * Generate an XLSX file as a `ReadableStream<Uint8Array>`.
 */
export function generateWorkbookStream(
  options: WorkbookOptions,
  packerOptions?: PackerOptions,
): ReadableStream<Uint8Array> {
  return Packer.toStream(options, packerOptions);
}
