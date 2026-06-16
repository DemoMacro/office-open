/**
 * Shared compiler utilities for OOXML document generation.
 *
 * @module
 */

import { levelForMediaName } from "../opc/packer";
import type { XmlifyedFile, ZipOptions, Zippable } from "../opc/packer";

/** Reusable TextEncoder instance (stateless, safe to share). */
const encoder = new TextEncoder();

/**
 * Convert a mapping of XML files + overrides + media into a Zippable structure.
 *
 * Centralises the mapping→Zippable conversion shared by all three compilers.
 */
export function compileMapping(
  mapping: Record<string, { data: string; path: string }>,
  overrides?: XmlifyedFile[],
  media?: { data: Uint8Array; path: string }[],
  mediaLevel: number = 0,
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
      // Already-compressed formats (PNG/JPEG/GIF) → STORE, else the media level.
      // Matches MS Office, which STORE-s pre-compressed assets and DEFLATE-s the rest.
      files[m.path] = [
        m.data,
        { level: levelForMediaName(m.path, mediaLevel) as ZipOptions["level"] },
      ];
    }
  }
  return files;
}
