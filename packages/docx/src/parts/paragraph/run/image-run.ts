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
import type { MediaData, NonVisualPropertiesOptions } from "@shared/media/data";

import type { Floating } from "../../drawing";
import type { GraphicFrameLocksOptions } from "../../drawing/descriptor";
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
  floating?: Floating;
  altText?: DocPropertiesOptions;
  outline?: OutlineOptions;
  fill?: FillOptions;
  effects?: EffectListOptions;
  blipEffects?: BlipEffectsOptions;
  sourceRectangle?: SourceRectangleOptions;
  tile?: TileOptions;
  /** Picture non-visual properties (pic:cNvPr) — populated by parse */
  nonVisualProperties?: NonVisualPropertiesOptions;
  /** Raw XML of the wrapping w:r's rPr (round-trip) — emitted before the drawing. */
  runPropertiesRawXml?: string;
  /** Graphic frame locks (wp:cNvGraphicFramePr) for round-trip. */
  graphicFrameLocks?: GraphicFrameLocksOptions | null;
  /** Blip rendering hint `a14:useLocalDpi` (round-trip). */
  useLocalDpi?: boolean;
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
  sourceRectangle?: SourceRectangleOptions,
  nonVisualProperties?: NonVisualPropertiesOptions,
): Pick<
  MediaData,
  "data" | "fileName" | "transformation" | "sourceRectangle" | "nonVisualProperties"
> => ({
  data,
  fileName: key,
  sourceRectangle,
  nonVisualProperties,
  transformation: createTransformation(transformation),
});
