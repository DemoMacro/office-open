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

/**
 * Options for color transforms.
 *
 * Percent fields take integer percent (e.g., `40` = 40%); the library scales
 * to the XSD 1/1000th-of-a-percent unit. Angle fields (`hue`/`hueOff`) take
 * degrees; the library scales to the XSD 1/60000th-of-a-degree unit.
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

/**
 * Creates color transform child elements as XML strings.
 *
 * These elements modify the parent color according to OOXML color transform rules.
 * Multiple transforms can be applied in sequence.
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

  if (options.tint !== undefined) t.push(`<a:tint val="${Math.round(options.tint * 1000)}"/>`);
  if (options.shade !== undefined) t.push(`<a:shade val="${Math.round(options.shade * 1000)}"/>`);
  if (options.comp) t.push(`<a:comp/>`);
  if (options.inv) t.push(`<a:inv/>`);
  if (options.gray) t.push(`<a:gray/>`);
  if (options.alpha !== undefined) t.push(`<a:alpha val="${Math.round(options.alpha * 1000)}"/>`);
  if (options.alphaOff !== undefined)
    t.push(`<a:alphaOff val="${Math.round(options.alphaOff * 1000)}"/>`);
  if (options.alphaMod !== undefined)
    t.push(`<a:alphaMod val="${Math.round(options.alphaMod * 1000)}"/>`);
  if (options.hue !== undefined) t.push(`<a:hue val="${Math.round(options.hue * 60000)}"/>`);
  if (options.hueOff !== undefined)
    t.push(`<a:hueOff val="${Math.round(options.hueOff * 60000)}"/>`);
  if (options.hueMod !== undefined)
    t.push(`<a:hueMod val="${Math.round(options.hueMod * 1000)}"/>`);
  if (options.sat !== undefined) t.push(`<a:sat val="${Math.round(options.sat * 1000)}"/>`);
  if (options.satOff !== undefined)
    t.push(`<a:satOff val="${Math.round(options.satOff * 1000)}"/>`);
  if (options.satMod !== undefined)
    t.push(`<a:satMod val="${Math.round(options.satMod * 1000)}"/>`);
  if (options.lum !== undefined) t.push(`<a:lum val="${Math.round(options.lum * 1000)}"/>`);
  if (options.lumOff !== undefined)
    t.push(`<a:lumOff val="${Math.round(options.lumOff * 1000)}"/>`);
  if (options.lumMod !== undefined)
    t.push(`<a:lumMod val="${Math.round(options.lumMod * 1000)}"/>`);
  if (options.red !== undefined) t.push(`<a:red val="${Math.round(options.red * 1000)}"/>`);
  if (options.redOff !== undefined)
    t.push(`<a:redOff val="${Math.round(options.redOff * 1000)}"/>`);
  if (options.redMod !== undefined)
    t.push(`<a:redMod val="${Math.round(options.redMod * 1000)}"/>`);
  if (options.green !== undefined) t.push(`<a:green val="${Math.round(options.green * 1000)}"/>`);
  if (options.greenOff !== undefined)
    t.push(`<a:greenOff val="${Math.round(options.greenOff * 1000)}"/>`);
  if (options.greenMod !== undefined)
    t.push(`<a:greenMod val="${Math.round(options.greenMod * 1000)}"/>`);
  if (options.blue !== undefined) t.push(`<a:blue val="${Math.round(options.blue * 1000)}"/>`);
  if (options.blueOff !== undefined)
    t.push(`<a:blueOff val="${Math.round(options.blueOff * 1000)}"/>`);
  if (options.blueMod !== undefined)
    t.push(`<a:blueMod val="${Math.round(options.blueMod * 1000)}"/>`);
  if (options.gamma) t.push(`<a:gamma/>`);
  if (options.invGamma) t.push(`<a:invGamma/>`);

  return t;
};
