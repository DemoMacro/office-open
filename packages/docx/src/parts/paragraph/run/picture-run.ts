import type { DataType } from "@office-open/core";
import type {
  BlipCompression,
  BlipEffectsOptions,
  EffectListOptions,
  FillOptions,
  OutlineOptions,
  Scene3DOptions,
  Shape3DOptions,
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
  /** 3D scene (pic:spPr/a:scene3d) — camera and lighting on the picture. */
  scene3d?: Scene3DOptions;
  /** 3D shape properties (pic:spPr/a:sp3d). */
  shape3d?: Shape3DOptions;
  blipEffects?: BlipEffectsOptions;
  sourceRectangle?: SourceRectangleOptions;
  tile?: TileOptions;
  /** Picture non-visual properties (pic:cNvPr) — populated by parse */
  nonVisualProperties?: NonVisualPropertiesOptions;
  /** Structured run properties of the wrapping w:r (round-trip) — emitted before the drawing. */
  runProperties?: RunPropertiesOptions;
  /** A w:lastRenderedPageBreak shared the drawing's run (round-trip) — emitted before the drawing. */
  lastRenderedPageBreak?: boolean;
  /** Graphic frame locks (wp:cNvGraphicFramePr) for round-trip. */
  graphicFrameLocks?: GraphicFrameLocksOptions | null;
  /** Blip rendering hint `a14:useLocalDpi` (round-trip). */
  useLocalDpi?: boolean;
  /**
   * External image source URL (a:blip @r:link) — BasePictureOptions.sourceUrl
   * spelled on the docx union (which does not extend the base). Paired with
   * data it is the linked source of the local cache; alone it is linked-only.
   */
  sourceUrl?: string;
  /** Compression state (a:blip `@cstate`); absent = attribute omitted (schema default "none"). */
  compression?: BlipCompression;
}

interface RegularPictureOptions {
  type: "jpg" | "png" | "gif" | "bmp" | "tif" | "ico" | "emf" | "wmf";
  /**
   * Image binary. Absent on a linked-only picture (external sourceUrl, no
   * bytes in the package) — the emitted a:blip carries r:link alone.
   */
  data?: DataType;
  /**
   * Source media file basename (round-trip). `type` normalizes jpeg→jpg etc.,
   * which would otherwise rewrite imageN.jpeg to imageN.jpg and drop the
   * source [Content_Types] Default extension. Omit for fresh authoring.
   */
  fileName?: string;
}

interface SvgMediaOptions {
  type: "svg";
  data: DataType;
  /**
   * Required in case the Word processor does not support SVG. A raster
   * fallback always carries bytes — the linked-only form has no vector part
   * to fall back from.
   */
  fallback: RegularPictureOptions & { data: DataType };
  /** Source vector file basename (round-trip). See RegularPictureOptions. */
  fileName?: string;
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
