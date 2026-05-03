import type { XmlComponent } from "@file/xml-components";
import {
    createGradientFill,
    createGroupFill,
    createNoFill,
    createPatternFill,
    createSolidFill,
} from "@office-open/core/drawingml";
import type { PatternFillOptions as CorePatternFillOptions } from "@office-open/core/drawingml";

/**
 * PPTX gradient stop options.
 * Position is 0-100 (percentage), color is a hex string like "FF0000" or "#FF0000".
 */
export interface GradientStopOptions {
    readonly position: number;
    readonly color: string;
}

/**
 * PPTX fill options — discriminated union.
 *
 * Supports string shorthand for solid fill (most common case).
 *
 * @example
 * // Solid fill (shorthand)
 * { fill: "4472C4" }
 * // Solid fill (explicit)
 * { fill: { type: "solid", color: "4472C4" } }
 * // No fill
 * { fill: { type: "none" } }
 * // Gradient fill
 * { fill: { type: "gradient", angle: 90, stops: [{ position: 0, color: "4472C4" }, { position: 100, color: "ED7D31" }] } }
 * // Pattern fill
 * { fill: { type: "pattern", pattern: "cross", foregroundColor: "FF0000" } }
 * // Group fill
 * { fill: { type: "group" } }
 */
export type FillOptions =
    | string
    | { readonly type: "solid"; readonly color: string }
    | { readonly type: "none" }
    | {
          readonly type: "gradient";
          readonly angle?: number;
          readonly scaled?: boolean;
          readonly stops: readonly GradientStopOptions[];
      }
    | {
          readonly type: "pattern";
          readonly pattern: string;
          readonly foregroundColor?: string;
          readonly backgroundColor?: string;
      }
    | { readonly type: "group" };

function normalizeColor(color: string): string {
    return color.replace("#", "");
}

/**
 * Builds a DrawingML fill XmlComponent from a FillOptions config.
 */
export function buildFill(options: FillOptions): XmlComponent {
    if (typeof options === "string") {
        return createSolidFill({ value: normalizeColor(options) });
    }

    switch (options.type) {
        case "solid":
            return createSolidFill({ value: normalizeColor(options.color) });
        case "none":
            return createNoFill();
        case "gradient":
            return createGradientFill({
                stops: options.stops.map((stop) => ({
                    position: stop.position * 1000,
                    color: { value: normalizeColor(stop.color) },
                })),
                ...(options.angle !== undefined && {
                    shade: { angle: options.angle * 60000, scaled: options.scaled ?? true },
                }),
            });
        case "pattern": {
            const coreOpts: CorePatternFillOptions = {
                pattern: options.pattern as CorePatternFillOptions["pattern"],
                ...(options.foregroundColor && {
                    foregroundColor: { value: normalizeColor(options.foregroundColor) },
                }),
                ...(options.backgroundColor && {
                    backgroundColor: { value: normalizeColor(options.backgroundColor) },
                }),
            };
            return createPatternFill(coreOpts);
        }
        case "group":
            return createGroupFill();
    }
}
