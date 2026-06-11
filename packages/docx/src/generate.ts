/**
 * Pure function API for generating DOCX files.
 *
 * @module
 */

import { createPacker, OoxmlMimeType } from "@office-open/core";
import type { OutputByType, OutputType, PackerOptions } from "@office-open/core";
import type { DocumentOptions } from "@parts/core-properties";

import { compileDocument } from "./compiler";

/** @internal Packer instance for DOCX generation. */
const Packer = createPacker<DocumentOptions>({
  compile: (options, overrides, mediaLevel) => compileDocument(options, overrides, mediaLevel),
  mimeType: OoxmlMimeType.DOCX,
});

/**
 * Generate a DOCX file from pure JSON options.
 *
 * The output format is controlled by `packerOptions.type` (default: `"nodebuffer"` → Buffer).
 * For synchronous generation, use {@link generateDocumentSync}. For streaming, use {@link generateDocumentStream}.
 *
 * @param options - Document options (sections, styles, numbering, etc.)
 * @param packerOptions - Optional packer configuration (type, compression, overrides, etc.)
 *
 * @example
 * ```typescript
 * import { generateDocument } from "@office-open/docx";
 *
 * const buffer = await generateDocument({ sections: [...] });
 * const bytes = await generateDocument({ sections: [...] }, { type: "uint8array" });
 * const blob = await generateDocument({ sections: [...] }, { type: "blob" });
 * ```
 */
export function generateDocument<T extends OutputType = "nodebuffer">(
  options: DocumentOptions,
  packerOptions?: PackerOptions<T>,
): Promise<OutputByType[T]> {
  return Packer.pack(options, packerOptions) as Promise<OutputByType[T]>;
}

/**
 * Synchronously generate a DOCX file from pure JSON options.
 */
export function generateDocumentSync<T extends OutputType = "nodebuffer">(
  options: DocumentOptions,
  packerOptions?: PackerOptions<T>,
): OutputByType[T] {
  return Packer.packSync(options, packerOptions) as OutputByType[T];
}

/**
 * Generate a DOCX file as a `ReadableStream<Uint8Array>`.
 */
export function generateDocumentStream(
  options: DocumentOptions,
  packerOptions?: PackerOptions,
): ReadableStream<Uint8Array> {
  return Packer.toStream(options, packerOptions);
}
