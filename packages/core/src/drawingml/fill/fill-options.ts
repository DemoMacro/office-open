import { element } from "@office-open/xml";

import { toUint8Array } from "../../util/data-type";
import type { DataType } from "../../util/data-type";
import { uniqueId } from "../../util/generators";
import { createBlipEffects } from "../blip/blip-effects";
import type { BlipEffectsOptions } from "../blip/blip-effects";
import { createSourceRectangle } from "../blip/source-rectangle";
import type { SourceRectangleOptions } from "../blip/source-rectangle";
import { createTileInfo } from "../blip/tile";
import type { TileOptions } from "../blip/tile";
import type { SolidFillOptions } from "../color/solid-fill";
import { createSolidFill } from "../color/solid-fill";
import { createGradientFill } from "./gradient-fill";
import type { GradientFillOptions } from "./gradient-fill";
import type { PatternFillOptions } from "./pattern-fill";
import { createPatternFill } from "./pattern-fill";

/**
 * Gradient stop options (simplified API).
 * Position is 0-100 (percentage), color is a hex string or SolidFillOptions.
 */
export interface GradientStopOptions {
  position: number;
  color: string | SolidFillOptions;
}

/**
 * Blip fill options (image fill) for DrawingML shapes.
 */
export interface BlipFillConfigOptions {
  /** Image data: raw bytes, ArrayBuffer, or a base64 data URL string. */
  data: DataType;
  /** Image type */
  imageType: "png" | "jpg" | "gif" | "bmp" | "tif" | "ico" | "emf" | "wmf";
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
 * fill: { type: "gradient", options: { stops: [...], shade: { angle: 5400000 } } }
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
      pattern: string;
      foregroundColor?: string | SolidFillOptions;
      backgroundColor?: string | SolidFillOptions;
    }
  | { type: "group" };

function normalizeColor(color: string | SolidFillOptions): SolidFillOptions {
  return typeof color === "string" ? { value: color.replace("#", "") } : color;
}

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
  if (typeof fill === "string" || fill.type !== "blip") return undefined;
  const raw = toUint8Array(fill.data);
  const fileName = nameAllocator
    ? nameAllocator(fill.imageType)
    : `${uniqueId()}.${fill.imageType}`;
  return { data: raw, fileName, type: fill.imageType };
};

/**
 * Builds a DrawingML fill XML string from a FillOptions config.
 */
export const buildFill = (options: FillOptions, embedPlaceholder?: string): string => {
  if (typeof options === "string") {
    return createSolidFill({ value: options.replace("#", "") });
  }

  switch (options.type) {
    case "solid":
      return createSolidFill(normalizeColor(options.color));
    case "none":
      return "<a:noFill/>";
    case "gradient":
      if ("options" in options) {
        return createGradientFill(options.options);
      }
      return createGradientFill({
        stops: options.stops.map((stop) => ({
          position: stop.position * 1000,
          color: normalizeColor(stop.color),
        })),
        ...(!options.path &&
          options.angle !== undefined && {
            shade: { angle: options.angle * 60000, scaled: options.scaled ?? true },
          }),
        ...(options.path && {
          shade: { path: options.path as "shape" | "circle" | "rect" },
        }),
      });
    case "blip": {
      const fileName = `${uniqueId()}.${options.imageType}`;

      // Build a:blip with {fileName} placeholder — the packer's ImageReplacer
      // will replace `{fileName}` with `rId{N}` and create the relationship.
      // We do NOT use createBlip here because it prefixes "rId" to the value,
      // which would produce "rIdrId{N}" after replacement. When the caller
      // supplies embedPlaceholder (a media reference already registered with
      // the write context, e.g. `{image1.png}`), use it verbatim instead of
      // minting a fresh id so the emitted reference matches the registration.
      const embed = embedPlaceholder ?? `{${fileName}}`;
      const blipChildren: string[] = [];
      if (options.blipEffects) {
        blipChildren.push(...createBlipEffects(options.blipEffects));
      }
      const blip = element(
        "a:blip",
        { cstate: "none", "r:embed": embed },
        blipChildren.length > 0 ? blipChildren : undefined,
      );

      const children: string[] = [blip, createSourceRectangle(options.sourceRectangle)];
      if (options.tile) {
        children.push(createTileInfo(options.tile));
      } else {
        children.push("<a:stretch><a:fillRect/></a:stretch>");
      }
      const attrs: Record<string, string | number | undefined> = {};
      if (options.dpi !== undefined) attrs.dpi = options.dpi;
      if (options.rotWithShape !== undefined) attrs.rotWithShape = options.rotWithShape ? 1 : 0;
      return element("a:blipFill", attrs, children);
    }
    case "pattern": {
      const coreOpts: PatternFillOptions = {
        pattern: options.pattern as PatternFillOptions["pattern"],
        ...(options.foregroundColor && {
          foregroundColor: normalizeColor(options.foregroundColor),
        }),
        ...(options.backgroundColor && {
          backgroundColor: normalizeColor(options.backgroundColor),
        }),
      };
      return createPatternFill(coreOpts);
    }
    case "group":
      return "<a:grpFill/>";
  }
};
