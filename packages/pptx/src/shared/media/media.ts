/**
 * Media module for PresentationML documents.
 *
 * Provides transformation helpers for embedded media. The deduplicated
 * {@link Media} collection lives in `@office`-open/core and is shared across all
 * format packages.
 *
 * @module
 */
import { convertEmuToPixels, convertToEmu } from "@office-open/core";
import type { UniversalMeasure } from "@office-open/core";

import type { MediaDataTransformation } from "./data";

export interface MediaTransformation {
  width: number | UniversalMeasure;
  height: number | UniversalMeasure;
}

export const createTransformation = (options: MediaTransformation): MediaDataTransformation => {
  const cx = convertToEmu(options.width);
  const cy = convertToEmu(options.height);
  return {
    emus: { x: cx, y: cy },
    pixels: { x: Math.round(convertEmuToPixels(cx)), y: Math.round(convertEmuToPixels(cy)) },
  };
};
