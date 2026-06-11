/**
 * Pure function API for generating PPTX files.
 *
 * @module
 */

import { createPacker, OoxmlMimeType } from "@office-open/core";
import type { OutputByType, OutputType, PackerOptions } from "@office-open/core";
import type { PresentationOptions } from "@shared/file";

import { compilePresentation } from "./compiler";

/** @internal Packer instance for PPTX generation. */
const Packer = createPacker<PresentationOptions>({
  compile: (options, overrides, mediaLevel) => compilePresentation(options, overrides, mediaLevel),
  mimeType: OoxmlMimeType.PPTX,
});

/**
 * Generate a PPTX file from pure JSON options.
 *
 * The output format is controlled by `packerOptions.type` (default: `"nodebuffer"` → Buffer).
 * For synchronous generation, use {@link generatePresentationSync}. For streaming, use {@link generatePresentationStream}.
 *
 * @param options - Presentation options (slides, masters, themes, etc.)
 * @param packerOptions - Optional packer configuration (type, compression, overrides, etc.)
 *
 * @example
 * ```typescript
 * import { generatePresentation } from "@office-open/pptx";
 *
 * const buffer = await generatePresentation({ slides: [...] });
 * const bytes = await generatePresentation({ slides: [...] }, { type: "uint8array" });
 * const blob = await generatePresentation({ slides: [...] }, { type: "blob" });
 * ```
 */
export function generatePresentation<T extends OutputType = "nodebuffer">(
  options: PresentationOptions,
  packerOptions?: PackerOptions<T>,
): Promise<OutputByType[T]> {
  return Packer.pack(options, packerOptions) as Promise<OutputByType[T]>;
}

/**
 * Synchronously generate a PPTX file from pure JSON options.
 */
export function generatePresentationSync<T extends OutputType = "nodebuffer">(
  options: PresentationOptions,
  packerOptions?: PackerOptions<T>,
): OutputByType[T] {
  return Packer.packSync(options, packerOptions) as OutputByType[T];
}

/**
 * Generate a PPTX file as a `ReadableStream<Uint8Array>`.
 */
export function generatePresentationStream(
  options: PresentationOptions,
  packerOptions?: PackerOptions,
): ReadableStream<Uint8Array> {
  return Packer.toStream(options, packerOptions);
}
