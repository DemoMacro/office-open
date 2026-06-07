/**
 * Pure function API for generating PPTX files.
 *
 * @module
 */

import type { PresentationOptions } from "@file/file";
import type { OutputByType, OutputType, PackerOptions } from "@office-open/core";

import { Packer } from "./packer/packer";

/**
 * Generate a PPTX file from pure JSON options.
 *
 * The output format is controlled by `packerOptions.type` (default: `"nodebuffer"` → Buffer).
 * For synchronous generation, use {@link generateSync}. For streaming, use {@link generateStream}.
 *
 * @param options - Presentation options (slides, masters, themes, etc.)
 * @param packerOptions - Optional packer configuration (type, compression, overrides, etc.)
 *
 * @example
 * ```typescript
 * import { generate } from "@office-open/pptx";
 *
 * const buffer = await generate({ slides: [...] });
 * const bytes = await generate({ slides: [...] }, { type: "uint8array" });
 * const blob = await generate({ slides: [...] }, { type: "blob" });
 * ```
 */
export function generate<T extends OutputType = "nodebuffer">(
  options: PresentationOptions,
  packerOptions?: PackerOptions<T>,
): Promise<OutputByType[T]> {
  return Packer.pack(options, packerOptions) as Promise<OutputByType[T]>;
}

/**
 * Synchronously generate a PPTX file from pure JSON options.
 */
export function generateSync<T extends OutputType = "nodebuffer">(
  options: PresentationOptions,
  packerOptions?: PackerOptions<T>,
): OutputByType[T] {
  return Packer.packSync(options, packerOptions) as OutputByType[T];
}

/**
 * Generate a PPTX file as a `ReadableStream<Uint8Array>`.
 */
export function generateStream(
  options: PresentationOptions,
  packerOptions?: PackerOptions,
): ReadableStream<Uint8Array> {
  return Packer.toStream(options, packerOptions);
}
