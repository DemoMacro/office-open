import { convertToEmu } from "@office-open/core";

/**
 * Media module for WordprocessingML documents.
 *
 * Provides transformation helpers for embedded media (images). The deduplicated
 * media collection itself ({@link Media}) lives in `@office`-open/core and is
 * re-exported from this package's index so every format package shares one
 * content-based dedup implementation.
 *
 * @module
 */
import type { MediaTransformation } from "./data";

/**
 * Normalizes authoring and parsed transformations to the same canonical shape.
 */
export const createTransformation = (options: MediaTransformation): MediaTransformation => {
  const width = convertToEmu(options.width);
  const height = convertToEmu(options.height);
  const x = convertToEmu(options.offset?.x ?? 0);
  const y = convertToEmu(options.offset?.y ?? 0);
  return {
    width,
    height,
    flipHorizontal: options.flipHorizontal,
    flipVertical: options.flipVertical,
    ...(options.offset ? { offset: { x, y } } : {}),
    rotation: options.rotation,
    ...(options.shapeExtent ? { shapeExtent: options.shapeExtent } : {}),
    ...(options.effectExtent ? { effectExtent: options.effectExtent } : {}),
    ...(options.inlineDistances ? { inlineDistances: options.inlineDistances } : {}),
  };
};
