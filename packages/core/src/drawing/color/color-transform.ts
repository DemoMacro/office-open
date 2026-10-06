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
import type { Percentage } from "../../util/values";

/**
 * Options for color transforms. Percent fields take integer percent (`40` =
 * 40%), angle fields (`hue`/`hueOff`) take degrees; the library applies the
 * XSD scaling. Booleans emit value-less switch elements.
 */
/** Integer percent (`50` = 50%) or a source ST_Percentage literal (`"50%"`). */
type TransformPercent = number | Percentage;

/**
 * Options for color transforms. Percent fields accept integer percent (`40` =
 * 40%) or a verbatim percent literal; angle fields take degrees. Booleans emit
 * value-less switch elements.
 */
export interface ColorTransformOptions {
  /** Tint: moves color toward white (0-100, where 100 = full white) */
  tint?: TransformPercent;
  /** Shade: moves color toward black (0-100, where 100 = full black) */
  shade?: TransformPercent;
  /** Complement: inverts the color (no value) */
  comp?: boolean;
  /** Inverse: inverts the color (no value) */
  inv?: boolean;
  /** Grayscale: converts to grayscale (no value) */
  gray?: boolean;
  /** Alpha: sets transparency (0-100, where 0 = transparent) */
  alpha?: TransformPercent;
  /** Alpha offset: adjusts alpha by fixed percent (-100 to 100) */
  alphaOff?: TransformPercent;
  /** Alpha modulation: scales alpha by percent (0-100) */
  alphaMod?: TransformPercent;
  /** Hue: sets hue angle in degrees (0-360). */
  hue?: number;
  /** Hue offset: adjusts hue angle in degrees (-90 to 90). */
  hueOff?: number;
  /** Hue modulation: scales hue by percent (0-100) */
  hueMod?: TransformPercent;
  /** Saturation: sets saturation (-100 to 100) */
  sat?: TransformPercent;
  /** Saturation offset: adjusts saturation (-100 to 100) */
  satOff?: TransformPercent;
  /** Saturation modulation: scales saturation (0-100) */
  satMod?: TransformPercent;
  /** Luminance: sets luminance (-100 to 100) */
  lum?: TransformPercent;
  /** Luminance offset: adjusts luminance (-100 to 100) */
  lumOff?: TransformPercent;
  /** Luminance modulation: scales luminance (0-100) */
  lumMod?: TransformPercent;
  /** Red: sets red channel (-100 to 100) */
  red?: TransformPercent;
  /** Red offset: adjusts red channel (-100 to 100) */
  redOff?: TransformPercent;
  /** Red modulation: scales red channel (0-100) */
  redMod?: TransformPercent;
  /** Green: sets green channel (-100 to 100) */
  green?: TransformPercent;
  /** Green offset: adjusts green channel (-100 to 100) */
  greenOff?: TransformPercent;
  /** Green modulation: scales green channel (0-100) */
  greenMod?: TransformPercent;
  /** Blue: sets blue channel (-100 to 100) */
  blue?: TransformPercent;
  /** Blue offset: adjusts blue channel (-100 to 100) */
  blueOff?: TransformPercent;
  /** Blue modulation: scales blue channel (0-100) */
  blueMod?: TransformPercent;
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
function emitTransformValue(key: TransformKey, value: number | Percentage): number | Percentage {
  if (typeof value === "string") return value;
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
