import { hashedId } from "../../id-generators";
import { BuilderElement, type XmlComponent } from "../../xml-components";
import { createBlipEffects } from "../blip/blip-effects";
import type { BlipEffectsOptions } from "../blip/blip-effects";
import { createSourceRectangle } from "../blip/source-rectangle";
import type { SourceRectangleOptions } from "../blip/source-rectangle";
import { Stretch } from "../blip/stretch";
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
    readonly position: number;
    readonly color: string | SolidFillOptions;
}

/**
 * Blip fill options (image fill) for DrawingML shapes.
 */
export interface BlipFillConfigOptions {
    /** Image data (raw bytes) */
    readonly data: Uint8Array | ArrayBuffer | Buffer;
    /** Image type */
    readonly imageType: "png" | "jpg" | "gif" | "bmp" | "tif" | "ico" | "emf" | "wmf";
    /** DPI of the image */
    readonly dpi?: number;
    /** Whether the fill rotates with the shape */
    readonly rotWithShape?: boolean;
    /** Image adjustment effects (brightness, contrast, grayscale, etc.) */
    readonly blipEffects?: BlipEffectsOptions;
    /** Source rectangle for cropping */
    readonly srcRect?: SourceRectangleOptions;
    /** Tile fill mode (if omitted, defaults to stretch) */
    readonly tile?: TileOptions;
}

/**
 * Media data extracted from a blip fill, for registration with the document's media store.
 */
export interface BlipFillMediaData {
    readonly fileName: string;
    readonly data: Uint8Array;
    readonly type: string;
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
    | { readonly type: "solid"; readonly color: string | SolidFillOptions }
    | { readonly type: "none" }
    | {
          readonly type: "gradient";
          readonly angle?: number;
          readonly scaled?: boolean;
          readonly path?: "shape" | "circle" | "rect";
          readonly stops: readonly GradientStopOptions[];
      }
    | { readonly type: "gradient"; readonly options: GradientFillOptions }
    | ({ readonly type: "blip" } & BlipFillConfigOptions)
    | {
          readonly type: "pattern";
          readonly pattern: string;
          readonly foregroundColor?: string | SolidFillOptions;
          readonly backgroundColor?: string | SolidFillOptions;
      }
    | { readonly type: "group" };

function normalizeColor(color: string | SolidFillOptions): SolidFillOptions {
    return typeof color === "string" ? { value: color.replace("#", "") } : color;
}

function toUint8Array(data: Uint8Array | ArrayBuffer | Buffer): Uint8Array {
    return data instanceof Uint8Array ? data : new Uint8Array(data);
}

/**
 * Extracts media data from a blip fill option, if present.
 * Returns undefined for non-blip fills.
 *
 * The returned data should be registered with the document's media store
 * during `prepForXml` so the packer can resolve the `{fileName}` placeholder.
 */
export const extractBlipFillMedia = (fill: FillOptions): BlipFillMediaData | undefined => {
    if (typeof fill === "string" || fill.type !== "blip") return undefined;
    const raw = toUint8Array(fill.data);
    const hash = hashedId(raw);
    const fileName = `${hash}.${fill.imageType}`;
    return { data: raw, fileName, type: fill.imageType };
};

/**
 * Builds a DrawingML fill XmlComponent from a FillOptions config.
 */
export const buildFill = (options: FillOptions): XmlComponent => {
    if (typeof options === "string") {
        return createSolidFill({ value: options.replace("#", "") });
    }

    switch (options.type) {
        case "solid":
            return createSolidFill(normalizeColor(options.color));
        case "none":
            return new BuilderElement({ name: "a:noFill" });
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
                    shade: { path: options.path.toUpperCase() as "SHAPE" | "CIRCLE" | "RECT" },
                }),
            });
        case "blip": {
            const raw = toUint8Array(options.data);
            const hash = hashedId(raw);
            const fileName = `${hash}.${options.imageType}`;

            // Build a:blip with {fileName} placeholder — the packer's ImageReplacer
            // will replace `{fileName}` with `rId{N}` and create the relationship.
            // We do NOT use createBlip here because it prefixes "rId" to the value,
            // which would produce "rIdrId{N}" after replacement.
            const blipChildren: XmlComponent[] = [];
            if (options.blipEffects) {
                blipChildren.push(...createBlipEffects(options.blipEffects));
            }
            const blip = new BuilderElement<{
                readonly embed: string;
                readonly cstate: string;
            }>({
                attributes: {
                    cstate: { key: "cstate", value: "none" },
                    embed: { key: "r:embed", value: `{${fileName}}` },
                },
                children: blipChildren,
                name: "a:blip",
            });

            const children: XmlComponent[] = [blip, createSourceRectangle(options.srcRect)];
            if (options.tile) {
                children.push(createTileInfo(options.tile));
            } else {
                children.push(new Stretch());
            }
            const attributes: Record<string, { readonly key: string; readonly value: number }> = {};
            if (options.dpi !== undefined) {
                attributes.dpi = { key: "dpi", value: options.dpi };
            }
            if (options.rotWithShape !== undefined) {
                attributes.rotWithShape = {
                    key: "rotWithShape",
                    value: options.rotWithShape ? 1 : 0,
                };
            }
            return new BuilderElement({
                attributes: Object.keys(attributes).length > 0 ? (attributes as never) : undefined,
                children,
                name: "a:blipFill",
            });
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
            return new BuilderElement({ name: "a:grpFill" });
    }
};
