/**
 * Pure function API for generating DOCX files.
 *
 * @module
 */

import type { PropertiesOptions } from "@file/core-properties";
import { File } from "@file/file";
import type { OutputByType, OutputType, PackerOptions } from "@office-open/core";

import { Packer } from "./packer/packer";

/**
 * Generate a DOCX file from pure JSON options.
 *
 * The output format is controlled by `packerOptions.type` (default: `"nodebuffer"` → Buffer).
 * For synchronous generation, use {@link generateSync}. For streaming, use {@link generateStream}.
 *
 * @param options - Document options (sections, styles, numbering, etc.)
 * @param packerOptions - Optional packer configuration (type, compression, overrides, etc.)
 *
 * @example
 * ```typescript
 * import { generate } from "@office-open/docx";
 *
 * const buffer = await generate({ sections: [...] });
 * const bytes = await generate({ sections: [...] }, { type: "uint8array" });
 * const blob = await generate({ sections: [...] }, { type: "blob" });
 * ```
 */
export function generate<T extends OutputType = "nodebuffer">(
  options: PropertiesOptions,
  packerOptions?: PackerOptions<T>,
): Promise<OutputByType[T]> {
  const doc = new File(options);
  return Packer.pack(doc, packerOptions) as Promise<OutputByType[T]>;
}

/**
 * Synchronously generate a DOCX file from pure JSON options.
 */
export function generateSync<T extends OutputType = "nodebuffer">(
  options: PropertiesOptions,
  packerOptions?: PackerOptions<T>,
): OutputByType[T] {
  const doc = new File(options);
  return Packer.packSync(doc, packerOptions) as OutputByType[T];
}

/**
 * Generate a DOCX file as a `ReadableStream<Uint8Array>`.
 */
export function generateStream(
  options: PropertiesOptions,
  packerOptions?: PackerOptions,
): ReadableStream<Uint8Array> {
  const doc = new File(options);
  return Packer.toStream(doc, packerOptions);
}
