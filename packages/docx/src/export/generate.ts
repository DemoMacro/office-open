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
 * Supports all output types via the `type` parameter (default: `"nodebuffer"` → Buffer).
 * For synchronous generation, use {@link generateSync}. For streaming, use {@link generateStream}.
 *
 * @param options - Document options (sections, styles, numbering, etc.)
 * @param type - Output format (default: `"nodebuffer"`)
 * @param packerOptions - Optional packer configuration (compression, overrides, etc.)
 *
 * @example
 * ```typescript
 * import { generate } from "@office-open/docx";
 *
 * // Default: Buffer
 * const buffer = await generate({ sections: [...] });
 *
 * // Specific output type
 * const bytes = await generate({ sections: [...] }, "uint8array");
 * const blob = await generate({ sections: [...] }, "blob");
 * const base64 = await generate({ sections: [...] }, "base64");
 *
 * // With packer options
 * const buffer = await generate({ sections: [...] }, "nodebuffer", {
 *   compression: { xml: 1 },
 * });
 * ```
 */
export function generate<T extends OutputType = "nodebuffer">(
  options: PropertiesOptions,
  type?: T,
  packerOptions?: PackerOptions,
): Promise<OutputByType[T]> {
  const doc = new File(options);
  return Packer.pack(doc, type ?? "nodebuffer", packerOptions) as Promise<OutputByType[T]>;
}

/**
 * Synchronously generate a DOCX file from pure JSON options.
 *
 * @param options - Document options (sections, styles, numbering, etc.)
 * @param type - Output format (default: `"nodebuffer"`)
 * @param packerOptions - Optional packer configuration (compression, overrides, etc.)
 *
 * @example
 * ```typescript
 * import { generateSync } from "@office-open/docx";
 *
 * const buffer = generateSync({ sections: [...] });
 * const bytes = generateSync({ sections: [...] }, "uint8array");
 * ```
 */
export function generateSync<T extends OutputType = "nodebuffer">(
  options: PropertiesOptions,
  type?: T,
  packerOptions?: PackerOptions,
): OutputByType[T] {
  const doc = new File(options);
  return Packer.packSync(doc, type ?? "nodebuffer", packerOptions) as OutputByType[T];
}

/**
 * Generate a DOCX file as a `ReadableStream<Uint8Array>`.
 *
 * Uses fflate's `AsyncZipDeflate` for non-blocking compression.
 * Works in both Node.js and browsers (Web Streams API).
 *
 * @param options - Document options (sections, styles, numbering, etc.)
 * @param packerOptions - Optional packer configuration (compression, overrides, etc.)
 *
 * @example
 * ```typescript
 * import { generateStream } from "@office-open/docx";
 *
 * const stream = generateStream({ sections: [...] });
 * const response = new Response(stream, {
 *   headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
 * });
 * ```
 */
export function generateStream(
  options: PropertiesOptions,
  packerOptions?: PackerOptions,
): ReadableStream<Uint8Array> {
  const doc = new File(options);
  return Packer.toStream(doc, packerOptions);
}
