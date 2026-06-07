/**
 * Media transformation utilities for DrawingML.
 *
 * Converts user-facing transformation options (pixels) to internal
 * transformation data (pixels + EMUs).
 *
 * @module
 */
import { convertPixelsToEmu } from "../../converters";

/**
 * Internal media data transformation with both pixel and EMU values.
 */
export interface MediaDataTransformation {
  offset?: {
    pixels: {
      x: number;
      y: number;
    };
    emus?: {
      x: number;
      y: number;
    };
  };
  pixels: {
    /** Width in pixels */
    x: number;
    /** Height in pixels */
    y: number;
  };
  /** Display dimensions in EMUs (English Metric Units) */
  emus: {
    /** Width in EMUs (1 inch = 914400 EMUs) */
    x: number;
    /** Height in EMUs (1 inch = 914400 EMUs) */
    y: number;
  };
  /** Optional flip transformations */
  flip?: {
    /** Whether to flip the image vertically */
    vertical?: boolean;
    /** Whether to flip the image horizontally */
    horizontal?: boolean;
  };
  /** Optional rotation angle in degrees */
  rotation?: number;
}

/**
 * Transformation options for media display.
 *
 * Specifies how an image should be transformed when displayed in the document.
 */
export interface MediaTransformation {
  offset?: {
    top?: number;
    left?: number;
  };
  width: number;
  /** Display height in pixels */
  height: number;
  /** Optional flip transformations */
  flip?: {
    /** Whether to flip the image vertically */
    vertical?: boolean;
    /** Whether to flip the image horizontally */
    horizontal?: boolean;
  };
  /** Optional rotation angle in degrees */
  rotation?: number;
}

/**
 * Converts user-facing transformation options (pixels) to internal
 * transformation data (pixels + EMUs).
 *
 * @param options - User-facing transformation in pixels
 * @returns Internal transformation data with both pixel and EMU values
 */
export const createTransformation = (options: MediaTransformation): MediaDataTransformation => ({
  emus: {
    x: convertPixelsToEmu(options.width),
    y: convertPixelsToEmu(options.height),
  },
  flip: options.flip,
  offset: {
    emus: {
      x: convertPixelsToEmu(options.offset?.left ?? 0),
      y: convertPixelsToEmu(options.offset?.top ?? 0),
    },
    pixels: {
      x: Math.round(options.offset?.left ?? 0),
      y: Math.round(options.offset?.top ?? 0),
    },
  },
  pixels: {
    x: Math.round(options.width),
    y: Math.round(options.height),
  },
  rotation: options.rotation ? options.rotation * 60_000 : undefined,
});
