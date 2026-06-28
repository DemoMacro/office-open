import type { UniversalMeasure } from "@office-open/core";
import { convertEmuToPixels, convertToEmu } from "@office-open/core";

/**
 * Media module for WordprocessingML documents.
 *
 * Provides transformation helpers for embedded media (images). The deduplicated
 * media collection itself ({@link Media}) lives in @office-open/core and is
 * re-exported from this package's index so every format package shares one
 * content-based dedup implementation.
 *
 * @module
 */
import type { MediaDataTransformation } from "./data";

/**
 * Transformation options for media display.
 *
 * Specifies how an image should be transformed when displayed in the document.
 * Width and height can be specified as numbers (in EMUs) or as universal measures
 * (e.g., "100mm", "2in").
 */
export interface MediaTransformation {
  offset?: {
    top?: number | UniversalMeasure;
    left?: number | UniversalMeasure;
  };
  width: number | UniversalMeasure;
  /** Display height in EMUs or universal measure */
  height: number | UniversalMeasure;
  /** Optional flip transformations */
  flip?: {
    /** Whether to flip the image vertically */
    vertical?: boolean;
    /** Whether to flip the image horizontally */
    horizontal?: boolean;
  };
  /** Optional rotation angle in degrees */
  rotation?: number;
  /** Effect extent (wp:effectExtent) in raw EMUs — passed through verbatim. */
  effectExtent?: { l: number; t: number; r: number; b: number };
}

/**
 * Converts user-facing transformation options (EMU or universal measure) to internal
 * transformation data (pixels + EMUs).
 *
 * @param options - User-facing transformation in EMU or universal measure
 * @returns Internal transformation data with both pixel and EMU values
 */
export const createTransformation = (options: MediaTransformation): MediaDataTransformation => {
  const widthEmu = convertToEmu(options.width);
  const heightEmu = convertToEmu(options.height);
  const offsetLeftEmu = convertToEmu(options.offset?.left ?? 0);
  const offsetTopEmu = convertToEmu(options.offset?.top ?? 0);
  return {
    emus: { x: widthEmu, y: heightEmu },
    flip: options.flip,
    offset: {
      emus: { x: offsetLeftEmu, y: offsetTopEmu },
      pixels: {
        x: Math.round(convertEmuToPixels(offsetLeftEmu)),
        y: Math.round(convertEmuToPixels(offsetTopEmu)),
      },
    },
    pixels: {
      x: Math.round(convertEmuToPixels(widthEmu)),
      y: Math.round(convertEmuToPixels(heightEmu)),
    },
    rotation: options.rotation ? options.rotation * 60_000 : undefined,
    ...(options.effectExtent ? { effectExtent: options.effectExtent } : {}),
  };
};
