import { toUint8Array } from "../../util/data-type";
import type { DataType } from "../../util/data-type";
import { uniqueId } from "../../util/generators";
import type { HexColor } from "../../util/values";
import type { BlipEffectsOptions } from "../blip/blip-effects";
import type { SourceRectangleOptions } from "../blip/source-rectangle";
import type { TileOptions } from "../blip/tile";
import type { SolidFillOptions } from "../color/solid-fill";
import { emitFillXml } from "./fill-descriptors";
import type { GradientFillOptions } from "./gradient-fill";
import type { PresetPattern } from "./pattern-fill";

/**
 * Gradient stop options (simplified API).
 * Position is 0-100 (percentage), color is a hex string or SolidFillOptions.
 */
export interface GradientStopOptions {
  position: number;
  color: HexColor | SolidFillOptions;
}

/**
 * Blip fill options (image fill) for DrawingML shapes.
 */
export interface BlipFillConfigOptions {
  /**
   * An empty a:blip carrying no r:embed — Word emits this in a pic:spPr
   * duplicate of pic:blipFill (the spPr copy references no image of its
   * own). No media is registered; the blip re-emits empty. When set, data
   * and imageType are meaningless.
   */
  noEmbed?: true;
  /** Image data: raw bytes, ArrayBuffer, or a base64 data URL string. */
  data?: DataType;
  imageType?: "png" | "jpg" | "gif" | "bmp" | "tif" | "ico" | "emf" | "wmf";
  /**
   * Source media file name, pinned on round-trip so re-emitting registers the
   * media under its original name instead of a re-derived one (jpeg→jpg
   * normalization would otherwise fork a second media part).
   */
  fileName?: string;
  /** DPI of the image */
  dpi?: number;
  /** Whether the fill rotates with the shape */
  rotWithShape?: boolean;
  /** Image adjustment effects (brightness, contrast, grayscale, etc.) */
  blipEffects?: BlipEffectsOptions;
  /** Source rectangle for cropping */
  sourceRectangle?: SourceRectangleOptions;
  /** Tile fill mode (if omitted, defaults to stretch) */
  tile?: TileOptions;
}

/**
 * Media data extracted from a blip fill, for registration with the document's media store.
 */
export interface BlipFillMediaData {
  fileName: string;
  data: Uint8Array;
  type: string;
}

/**
 * Fill options — discriminated union for DrawingML EG_FillProperties.
 *
 * Supports string shorthand for solid fill (most common case).
 * Color fields accept hex strings or advanced SolidFillOptions
 * (scheme color, HSL, etc.).
 *
 * @example
 * // Solid fill (shorthand)
 * fill: "4472C4"
 * // Solid fill with scheme color
 * fill: { type: "solid", color: { value: "accent1" } }
 * // No fill
 * fill: { type: "none" }
 * // Gradient fill — linear
 * fill: { type: "gradient", angle: 90, stops: [{ position: 0, color: "4472C4" }, { position: 100, color: "ED7D31" }] }
 * // Gradient fill — path (radial)
 * fill: { type: "gradient", path: "circle", stops: [{ position: 0, color: "FFFFFF" }, { position: 100, color: "4472C4" }] }
 * // Gradient fill (core API for advanced options)
 * fill: { type: "gradient", options: { stops: [...], shade: { angle: 90 } } }
 * // Blip fill (image)
 * fill: { type: "blip", data: imageBuffer, imageType: "png" }
 * // Pattern fill
 * fill: { type: "pattern", pattern: "cross", foregroundColor: "FF0000" }
 * // Group fill
 * fill: { type: "group" }
 */
export type FillOptions =
  | string
  | { type: "solid"; color: string | SolidFillOptions }
  | { type: "none" }
  | {
      type: "gradient";
      angle?: number;
      scaled?: boolean;
      path?: "shape" | "circle" | "rect";
      stops: readonly GradientStopOptions[];
    }
  | { type: "gradient"; options: GradientFillOptions }
  | ({ type: "blip" } & BlipFillConfigOptions)
  | {
      type: "pattern";
      /** Preset pattern (a:pattFill `@prst`, ST_PresetPatternVal). */
      pattern: PresetPattern;
      foregroundColor?: HexColor | SolidFillOptions;
      backgroundColor?: HexColor | SolidFillOptions;
    }
  | { type: "group" };

/**
 * Extracts media data from a blip fill option, if present.
 * Returns undefined for non-blip fills.
 *
 * The returned data should be registered with the document's media store
 * during serialization so the packer can resolve the `{fileName}` placeholder.
 *
 * @param fill - Fill options to inspect
 * @param nameAllocator - Optional sequential name provider (e.g. a format
 *   package's media counter). When omitted, falls back to a random id so the
 *   function stays usable from contexts without a shared counter.
 */
export const extractBlipFillMedia = (
  fill: FillOptions,
  nameAllocator?: (type: string) => string,
): BlipFillMediaData | undefined => {
  // A noEmbed blip fill carries no media of its own (empty a:blip marker).
  if (typeof fill === "string" || fill.type !== "blip" || !fill.data || !fill.imageType)
    return undefined;
  const raw = toUint8Array(fill.data);
  const fileName = nameAllocator
    ? nameAllocator(fill.imageType)
    : `${uniqueId()}.${fill.imageType}`;
  return { data: raw, fileName, type: fill.imageType };
};

/**
 * Builds a DrawingML fill XML string from a FillOptions config.
 * Thin delegate to the single serializer in the fill descriptors — kept as the
 * context-free public entry (blip fills mint a `{fileName}` embed placeholder
 * unless the caller supplies one).
 */
export const buildFill = (options: FillOptions, embedPlaceholder?: string): string | undefined =>
  emitFillXml(options, embedPlaceholder);
