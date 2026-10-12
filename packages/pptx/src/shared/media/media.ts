/**
 * Media module for PresentationML documents.
 *
 * Provides transformation helpers for embedded media. The deduplicated
 * {@link Media} collection lives in `@office`-open/core and is shared across all
 * format packages.
 *
 * @module
 */
import { convertToEmu } from "@office-open/core";

import type { MediaTransformation } from "./data";

export const createTransformation = (options: MediaTransformation): MediaTransformation => ({
  width: convertToEmu(options.width),
  height: convertToEmu(options.height),
});
