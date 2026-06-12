import type { DataType } from "@office-open/core";
import type { FillOptions } from "@office-open/core/drawingml";
/**
 * ImageRun types for WordprocessingML documents.
 *
 * This module provides support for inserting images into documents.
 *
 * Reference: http://officeopenxml.com/drwPicInline.php
 *
 * @module
 */
import type { DocPropertiesOptions } from "@parts/drawing/doc-properties/doc-properties";
import type { MediaTransformation } from "@shared/media";
import { createTransformation } from "@shared/media";
import type { MediaData } from "@shared/media/data";

import type { BlipEffectsOptions } from "../../drawing/inline/graphic/graphic-data/pic/blip/blip-effects";
import type { SourceRectangleOptions } from "../../drawing/inline/graphic/graphic-data/pic/blip/source-rectangle";
import type { TileOptions } from "../../drawing/inline/graphic/graphic-data/pic/blip/tile";
import type { EffectListOptions } from "../../drawing/inline/graphic/graphic-data/pic/effects/effect-list";
import type { OutlineOptions } from "../../drawing/inline/graphic/graphic-data/pic/outline/outline";

/**
 * Core options for image configuration.
 */
interface CoreImageOptions {
  transformation: MediaTransformation;
  floating?: import("../../drawing").Floating;
  altText?: DocPropertiesOptions;
  outline?: OutlineOptions;
  fill?: FillOptions;
  effects?: EffectListOptions;
  blipEffects?: BlipEffectsOptions;
  srcRect?: SourceRectangleOptions;
  tile?: TileOptions;
}

interface RegularImageOptions {
  type: "jpg" | "png" | "gif" | "bmp" | "tif" | "ico" | "emf" | "wmf";
  data: DataType;
}

interface SvgMediaOptions {
  type: "svg";
  data: DataType;
  /**
   * Required in case the Word processor does not support SVG.
   */
  fallback: RegularImageOptions;
}

/**
 * Options for creating an ImageRun.
 *
 * @see {@link ImageRun}
 */
export type ImageOptions = (RegularImageOptions | SvgMediaOptions) & CoreImageOptions;

export const createImageData = (
  data: Uint8Array,
  transformation: MediaTransformation,
  key: string,
  srcRect?: SourceRectangleOptions,
): Pick<MediaData, "data" | "fileName" | "transformation" | "srcRect"> => ({
  data,
  fileName: key,
  srcRect,
  transformation: createTransformation(transformation),
});
