/**
 * Pure function API for generating PPTX files.
 *
 * @module
 */

import type { PresentationOptions } from "@file/file";
import { File } from "@file/file";
import type { OutputByType, OutputType, PackerOptions } from "@office-open/core";

import { Packer } from "./packer/packer";

/**
 * Generate a PPTX file from pure JSON options.
 *
 * Supports all output types via the `type` parameter (default: `"nodebuffer"` → Buffer).
 * For synchronous generation, use {@link generateSync}. For streaming, use {@link generateStream}.
 *
 * @param options - Presentation options (slides, masters, theme, etc.)
 * @param type - Output format (default: `"nodebuffer"`)
 * @param packerOptions - Optional packer configuration (compression, overrides, etc.)
 *
 * @example
 * ```typescript
 * import { generate } from "@office-open/pptx";
 *
 * const buffer = await generate({ slides: [...] });
 * const bytes = await generate({ slides: [...] }, "uint8array");
 * const blob = await generate({ slides: [...] }, "blob");
 * ```
 */
export function generate<T extends OutputType = "nodebuffer">(
  options: PresentationOptions,
  type?: T,
  packerOptions?: PackerOptions,
): Promise<OutputByType[T]> {
  const pres = new File(options);
  return Packer.pack(pres, type ?? "nodebuffer", packerOptions) as Promise<OutputByType[T]>;
}

/**
 * Synchronously generate a PPTX file from pure JSON options.
 */
export function generateSync<T extends OutputType = "nodebuffer">(
  options: PresentationOptions,
  type?: T,
  packerOptions?: PackerOptions,
): OutputByType[T] {
  const pres = new File(options);
  return Packer.packSync(pres, type ?? "nodebuffer", packerOptions) as OutputByType[T];
}

/**
 * Generate a PPTX file as a `ReadableStream<Uint8Array>`.
 */
export function generateStream(
  options: PresentationOptions,
  packerOptions?: PackerOptions,
): ReadableStream<Uint8Array> {
  const pres = new File(options);
  return Packer.toStream(pres, packerOptions);
}
