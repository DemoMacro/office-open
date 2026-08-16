/**
 * Media path → picture type token.
 *
 * @module
 */

import { imageTypeFromPath as coreImageTypeFromPath } from "@office-open/core";

import type { PictureOptions } from "../picture";

/**
 * Map a media path's file extension to a picture type token (unknown → png).
 * PPTX pictures have no tif/ico variant — those extensions fall back to png.
 */
export function imageTypeFromPath(path: string): PictureOptions["type"] {
  const type = coreImageTypeFromPath(path);
  return type === "tif" || type === "ico" ? "png" : type;
}
