/**
 * Shared compiler utilities for OOXML document generation.
 *
 * @module
 */

import type { XmlifyedFile, Zippable } from "./packer";

/** Default STORE level for already-compressed media formats. */
const ZIP_STORED_LEVEL = 0;

/** Reusable TextEncoder instance (stateless, safe to share). */
const encoder = new TextEncoder();

/**
 * Convert a mapping of XML files + overrides + media into a Zippable structure.
 *
 * Centralises the mapping→Zippable conversion shared by all three compilers.
 */
export function compileMapping(
  mapping: Record<string, { data: string; path: string }>,
  overrides?: readonly XmlifyedFile[],
  media?: readonly { data: Uint8Array; path: string }[],
): Zippable {
  const files: Zippable = {};
  for (const entry of Object.values(mapping)) {
    files[entry.path] = encoder.encode(entry.data);
  }
  if (overrides) {
    for (const o of overrides) {
      files[o.path] = typeof o.data === "string" ? encoder.encode(o.data) : o.data;
    }
  }
  if (media) {
    for (const m of media) {
      files[m.path] = [m.data, { level: ZIP_STORED_LEVEL }];
    }
  }
  return files;
}
