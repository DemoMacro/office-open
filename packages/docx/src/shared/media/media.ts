import type { UniversalMeasure } from "@office-open/core";
import { convertPixelsToEmu, convertUniversalMeasureToEmu } from "@office-open/core";

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
 * Width and height can be specified as numbers (in pixels) or as universal measures
 * (e.g., "100mm", "2in").
 */
export interface MediaTransformation {
  offset?: {
    top?: number | UniversalMeasure;
    left?: number | UniversalMeasure;
  };
  width: number | UniversalMeasure;
  /** Display height in pixels or universal measure */
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
 * Converts user-facing transformation options (pixels or universal measure) to internal
 * transformation data (pixels + EMUs).
 *
 * @param options - User-facing transformation in pixels or universal measure
 * @returns Internal transformation data with both pixel and EMU values
 */
export const createTransformation = (options: MediaTransformation): MediaDataTransformation => ({
  emus: {
    x:
      typeof options.width === "string"
        ? convertUniversalMeasureToEmu(options.width)
        : convertPixelsToEmu(options.width),
    y:
      typeof options.height === "string"
        ? convertUniversalMeasureToEmu(options.height)
        : convertPixelsToEmu(options.height),
  },
  flip: options.flip,
  offset: {
    emus: {
      x:
        typeof options.offset?.left === "string"
          ? convertUniversalMeasureToEmu(options.offset.left)
          : convertPixelsToEmu(options.offset?.left ?? 0),
      y:
        typeof options.offset?.top === "string"
          ? convertUniversalMeasureToEmu(options.offset.top)
          : convertPixelsToEmu(options.offset?.top ?? 0),
    },
    pixels: {
      x: typeof options.offset?.left === "number" ? Math.round(options.offset.left) : 0,
      y: typeof options.offset?.top === "number" ? Math.round(options.offset.top) : 0,
    },
  },
  pixels: {
    x: typeof options.width === "number" ? Math.round(options.width) : 0,
    y: typeof options.height === "number" ? Math.round(options.height) : 0,
  },
  rotation: options.rotation ? options.rotation * 60_000 : undefined,
  ...(options.effectExtent ? { effectExtent: options.effectExtent } : {}),
});
