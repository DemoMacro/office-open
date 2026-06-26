/**
 * Media module for PresentationML documents.
 *
 * Provides transformation helpers for embedded media. The deduplicated
 * {@link Media} collection lives in @office-open/core and is shared across all
 * format packages.
 *
 * @module
 */
import { convertPixelsToEmu } from "@office-open/core";

import type { MediaDataTransformation } from "./data";

export interface MediaTransformation {
  offset?: {
    top?: number;
    left?: number;
  };
  width: number;
  height: number;
  flip?: {
    vertical?: boolean;
    horizontal?: boolean;
  };
  rotation?: number;
}

export const createTransformation = (options: MediaTransformation): MediaDataTransformation => ({
  emus: {
    x: convertPixelsToEmu(options.width),
    y: convertPixelsToEmu(options.height),
  },
  pixels: {
    x: Math.round(options.width),
    y: Math.round(options.height),
  },
  flip: options.flip,
  rotation: options.rotation ? options.rotation * 60_000 : undefined,
});
