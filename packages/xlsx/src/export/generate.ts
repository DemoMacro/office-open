/**
 * Pure function API for generating XLSX files.
 *
 * @module
 */

import type { WorkbookOptions } from "@file/file";
import { File } from "@file/file";
import type { OutputByType, OutputType, PackerOptions } from "@office-open/core";

import { Packer } from "./packer/packer";

/**
 * Generate an XLSX file from pure JSON options.
 *
 * Supports all output types via the `type` parameter (default: `"nodebuffer"` → Buffer).
 * For synchronous generation, use {@link generateSync}. For streaming, use {@link generateStream}.
 *
 * @param options - Workbook options (worksheets, styles, etc.)
 * @param type - Output format (default: `"nodebuffer"`)
 * @param packerOptions - Optional packer configuration (compression, overrides, etc.)
 *
 * @example
 * ```typescript
 * import { generate } from "@office-open/xlsx";
 *
 * const buffer = await generate({ worksheets: [...] });
 * const bytes = await generate({ worksheets: [...] }, "uint8array");
 * const blob = await generate({ worksheets: [...] }, "blob");
 * ```
 */
export function generate<T extends OutputType = "nodebuffer">(
  options: WorkbookOptions,
  type?: T,
  packerOptions?: PackerOptions,
): Promise<OutputByType[T]> {
  const workbook = new File(options);
  return Packer.pack(workbook, type ?? "nodebuffer", packerOptions) as Promise<OutputByType[T]>;
}

/**
 * Synchronously generate an XLSX file from pure JSON options.
 */
export function generateSync<T extends OutputType = "nodebuffer">(
  options: WorkbookOptions,
  type?: T,
  packerOptions?: PackerOptions,
): OutputByType[T] {
  const workbook = new File(options);
  return Packer.packSync(workbook, type ?? "nodebuffer", packerOptions) as OutputByType[T];
}

/**
 * Generate an XLSX file as a `ReadableStream<Uint8Array>`.
 */
export function generateStream(
  options: WorkbookOptions,
  packerOptions?: PackerOptions,
): ReadableStream<Uint8Array> {
  const workbook = new File(options);
  return Packer.toStream(workbook, packerOptions);
}
