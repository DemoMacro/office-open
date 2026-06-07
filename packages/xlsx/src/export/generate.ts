/**
 * Pure function API for generating XLSX files.
 *
 * @module
 */

import type { WorkbookOptions } from "@file/file";
import type { OutputByType, OutputType, PackerOptions } from "@office-open/core";

import { Packer } from "./packer/packer";

/**
 * Generate an XLSX file from pure JSON options.
 *
 * The output format is controlled by `packerOptions.type` (default: `"nodebuffer"` → Buffer).
 * For synchronous generation, use {@link generateSync}. For streaming, use {@link generateStream}.
 *
 * @param options - Workbook options (worksheets, styles, etc.)
 * @param packerOptions - Optional packer configuration (type, compression, overrides, etc.)
 *
 * @example
 * ```typescript
 * import { generate } from "@office-open/xlsx";
 *
 * const buffer = await generate({ worksheets: [...] });
 * const bytes = await generate({ worksheets: [...] }, { type: "uint8array" });
 * const blob = await generate({ worksheets: [...] }, { type: "blob" });
 * ```
 */
export function generate<T extends OutputType = "nodebuffer">(
  options: WorkbookOptions,
  packerOptions?: PackerOptions<T>,
): Promise<OutputByType[T]> {
  return Packer.pack(options, packerOptions) as Promise<OutputByType[T]>;
}

/**
 * Synchronously generate an XLSX file from pure JSON options.
 */
export function generateSync<T extends OutputType = "nodebuffer">(
  options: WorkbookOptions,
  packerOptions?: PackerOptions<T>,
): OutputByType[T] {
  return Packer.packSync(options, packerOptions) as OutputByType[T];
}

/**
 * Generate an XLSX file as a `ReadableStream<Uint8Array>`.
 */
export function generateStream(
  options: WorkbookOptions,
  packerOptions?: PackerOptions,
): ReadableStream<Uint8Array> {
  return Packer.toStream(options, packerOptions);
}
