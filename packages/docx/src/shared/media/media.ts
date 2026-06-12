import type { UniversalMeasure } from "@office-open/core";
import { convertPixelsToEmu, convertUniversalMeasureToEmu } from "@office-open/core";

/**
 * Media module for WordprocessingML documents.
 *
 * This module provides support for managing embedded media (images)
 * within a document.
 *
 * @module
 */
import type { MediaDataTransformation } from "./data";
import type { MediaData } from "./data";

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
});

/**
 * Manages embedded media (images) in a document.
 *
 * Media stores all images referenced in the document and provides
 * access to their data for packaging into the DOCX file. Each image
 * is stored with a unique key for retrieval.
 *
 * @example
 * ```typescript
 * const media = new Media();
 * media.addImage("image1", {
 *   type: "png",
 *   fileName: "image1.png",
 *   transformation: {
 *     pixels: { x: 200, y: 100 },
 *     emus: { x: 1828800, y: 914400 }
 *   },
 *   data: imageBuffer
 * });
 * const allImages = media.Array;
 * ```
 */
export class Media {
  private map: Map<string, MediaData>;

  public constructor() {
    this.map = new Map<string, MediaData>();
  }

  /**
   * Adds an image to the media collection.
   *
   * @param key - Unique identifier for this image
   * @param mediaData - Complete image data including file name, transformation, and raw data
   */
  public addImage(key: string, mediaData: MediaData): void {
    this.map.set(key, mediaData);
  }

  /**
   * Gets all images as an array.
   *
   * @returns Read-only array of all media data in the collection
   */
  public get array(): MediaData[] {
    return [...this.map.values()];
  }
}
