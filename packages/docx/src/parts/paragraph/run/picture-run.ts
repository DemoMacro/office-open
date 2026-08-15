import type { DataType } from "@office-open/core";
import type {
  BlipEffectsOptions,
  EffectListOptions,
  FillOptions,
  OutlineOptions,
  SourceRectangleOptions,
  TileOptions,
} from "@office-open/core/drawing";
/**
 * Picture (pic:pic) run types for WordprocessingML documents.
 *
 * This module provides support for inserting pictures into documents.
 *
 * Reference: http://officeopenxml.com/drwPicInline.php
 *
 * @module
 */
import type { DocPropertiesOptions } from "@parts/drawing/doc-properties/doc-properties";
import type { RunPropertiesOptions } from "@parts/paragraph/run/properties";
import type { MediaTransformation } from "@shared/media";
import { createTransformation } from "@shared/media";
import type { MediaData, NonVisualPropertiesOptions } from "@shared/media/data";

import type { Floating } from "../../drawing";
import type { GraphicFrameLocksOptions } from "../../drawing/descriptor";

/**
 * Core options for picture configuration.
 */
interface CorePictureOptions {
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
  /** Structured run properties of the wrapping w:r (round-trip) — emitted before the drawing. */
  runProperties?: RunPropertiesOptions;
  /** Graphic frame locks (wp:cNvGraphicFramePr) for round-trip. */
  graphicFrameLocks?: GraphicFrameLocksOptions | null;
  /** Blip rendering hint `a14:useLocalDpi` (round-trip). */
  useLocalDpi?: boolean;
}

interface RegularPictureOptions {
  type: "jpg" | "png" | "gif" | "bmp" | "tif" | "ico" | "emf" | "wmf";
  data: DataType;
}

interface SvgMediaOptions {
  type: "svg";
  data: DataType;
  /**
   * Required in case the Word processor does not support SVG.
   */
  fallback: RegularPictureOptions;
}

/**
 * Options for an inline/anchored picture (pic:pic).
 */
export type PictureOptions = (RegularPictureOptions | SvgMediaOptions) & CorePictureOptions;

export const createPictureData = (
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
