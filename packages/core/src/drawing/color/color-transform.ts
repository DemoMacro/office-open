/**
 * Color transform elements for DrawingML colors.
 *
 * This module provides color transformation elements defined in EG_ColorTransform,
 * which can be applied as child elements to any color type (srgbClr, schemeClr, etc.).
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, EG_ColorTransform
 *
 * @module
 */

import { emitAngle, emitPercent } from "../../util/converters";

/**
 * Options for color transforms.
 *
 * Percent fields take integer percent (e.g., `40` = 40%); the library scales
 * to the XSD 1/1000th-of-a-percent unit. Angle fields (`hue`/`hueOff`) take
 * degrees; the library scales to the XSD 1/60000th-of-a-degree unit. Boolean
 * fields emit value-less switch elements.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, EG_ColorTransform
 */
export interface ColorTransformOptions {
  /** Tint: moves color toward white (0-100, where 100 = full white) */
  tint?: number;
  /** Shade: moves color toward black (0-100, where 100 = full black) */
  shade?: number;
  /** Complement: inverts the color (no value) */
  comp?: boolean;
  /** Inverse: inverts the color (no value) */
  inv?: boolean;
  /** Grayscale: converts to grayscale (no value) */
  gray?: boolean;
  /** Alpha: sets transparency (0-100, where 0 = transparent) */
  alpha?: number;
  /** Alpha offset: adjusts alpha by fixed percent (-100 to 100) */
  alphaOff?: number;
  /** Alpha modulation: scales alpha by percent (0-100) */
  alphaMod?: number;
  /** Hue: sets hue angle in degrees (0-360). */
  hue?: number;
  /** Hue offset: adjusts hue angle in degrees (-90 to 90). */
  hueOff?: number;
  /** Hue modulation: scales hue by percent (0-100) */
  hueMod?: number;
  /** Saturation: sets saturation (-100 to 100) */
  sat?: number;
  /** Saturation offset: adjusts saturation (-100 to 100) */
  satOff?: number;
  /** Saturation modulation: scales saturation (0-100) */
  satMod?: number;
  /** Luminance: sets luminance (-100 to 100) */
  lum?: number;
  /** Luminance offset: adjusts luminance (-100 to 100) */
  lumOff?: number;
  /** Luminance modulation: scales luminance (0-100) */
  lumMod?: number;
  /** Red: sets red channel (-100 to 100) */
  red?: number;
  /** Red offset: adjusts red channel (-100 to 100) */
  redOff?: number;
  /** Red modulation: scales red channel (0-100) */
  redMod?: number;
  /** Green: sets green channel (-100 to 100) */
  green?: number;
  /** Green offset: adjusts green channel (-100 to 100) */
  greenOff?: number;
  /** Green modulation: scales green channel (0-100) */
  greenMod?: number;
  /** Blue: sets blue channel (-100 to 100) */
  blue?: number;
  /** Blue offset: adjusts blue channel (-100 to 100) */
  blueOff?: number;
  /** Blue modulation: scales blue channel (0-100) */
  blueMod?: number;
  /** Gamma correction (no value) */
  gamma?: boolean;
  /** Inverse gamma correction (no value) */
  invGamma?: boolean;
}

type TransformKey = keyof ColorTransformOptions & string;

/**
 * Transform keys classified by XSD unit — the single source of truth shared by
 * stringify and parse. Percent keys take integer percent (e.g. `50` = 50%,
 * scaled ×1000); angle keys take degrees (scaled ×60000); value-less boolean
 * keys (comp/inv/gray/gamma/invGamma) belong to neither set.
 */
export const PERCENT_TRANSFORMS: ReadonlySet<TransformKey> = new Set<TransformKey>([
  "tint",
  "shade",
  "alpha",
  "alphaOff",
  "alphaMod",
  "hueMod",
  "sat",
  "satOff",
  "satMod",
  "lum",
  "lumOff",
  "lumMod",
  "red",
  "redOff",
  "redMod",
  "green",
  "greenOff",
  "greenMod",
  "blue",
  "blueOff",
  "blueMod",
]);

export const ANGLE_TRANSFORMS: ReadonlySet<TransformKey> = new Set<TransformKey>(["hue", "hueOff"]);

/** Value-less switch transforms (empty elements — presence is the semantics). */
export const BOOLEAN_TRANSFORMS: ReadonlySet<TransformKey> = new Set<TransformKey>([
  "comp",
  "inv",
  "gray",
  "gamma",
  "invGamma",
]);

/** Scale a transform value to its XSD unit; non-percent/angle keys pass through. */
export function emitTransformValue(key: TransformKey, value: number): number {
  if (PERCENT_TRANSFORMS.has(key)) return emitPercent(value);
  if (ANGLE_TRANSFORMS.has(key)) return emitAngle(value);
  return value;
}

/**
 * Serialize color transforms, preserving the key order of the options object.
 *
 * EG_ColorTransform is an unordered XSD choice, but the sequence is
 * semantically significant (transforms compose left to right), so the caller's
 * key order is the emission order and round-trips come back in source order —
 * plain JS objects preserve insertion order.
 *
 * @example
 * ```typescript
 * // Lighten accent1 by 40%
 * createColorTransforms({ tint: 40 });
 * // Semi-transparent red with 50% alpha
 * createColorTransforms({ alpha: 50 });
 * ```
 */
export const createColorTransforms = (options: ColorTransformOptions): readonly string[] => {
  const t: string[] = [];

  for (const [name, value] of Object.entries(options)) {
    const key = name as TransformKey;
    if (value === undefined || value === false) continue;
    if (value === true) {
      t.push(`<a:${key}/>`);
      continue;
    }
    t.push(`<a:${key} val="${emitTransformValue(key, value)}"/>`);
  }

  return t;
};
