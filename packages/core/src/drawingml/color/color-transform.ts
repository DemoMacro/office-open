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
 * All percentage values are in 1/1000th of a percent (e.g., 50000 = 50%).
 * Angle values are in 60,000ths of a degree.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, EG_ColorTransform
 */
export interface ColorTransformOptions {
  /** Tint: moves color toward white (0-100000) */
  tint?: number;
  /** Shade: moves color toward black (0-100000) */
  shade?: number;
  /** Complement: inverts the color (no value) */
  comp?: boolean;
  /** Inverse: inverts the color (no value) */
  inv?: boolean;
  /** Grayscale: converts to grayscale (no value) */
  gray?: boolean;
  /** Alpha: sets transparency (0-100000) */
  alpha?: number;
  /** Alpha offset: adjusts alpha by fixed amount (-100000 to 100000) */
  alphaOff?: number;
  /** Alpha modulation: scales alpha by percentage (0-100000) */
  alphaMod?: number;
  /** Hue: sets hue angle (0-21600000) */
  hue?: number;
  /** Hue offset: adjusts hue by angle (-5400000 to 5400000) */
  hueOff?: number;
  /** Hue modulation: scales hue by percentage (0-100000) */
  hueMod?: number;
  /** Saturation: sets saturation (-100000 to 100000) */
  sat?: number;
  /** Saturation offset: adjusts saturation (-100000 to 100000) */
  satOff?: number;
  /** Saturation modulation: scales saturation (0-100000) */
  satMod?: number;
  /** Luminance: sets luminance (-100000 to 100000) */
  lum?: number;
  /** Luminance offset: adjusts luminance (-100000 to 100000) */
  lumOff?: number;
  /** Luminance modulation: scales luminance (0-100000) */
  lumMod?: number;
  /** Red: sets red channel (-100000 to 100000) */
  red?: number;
  /** Red offset: adjusts red channel (-100000 to 100000) */
  redOff?: number;
  /** Red modulation: scales red channel (0-100000) */
  redMod?: number;
  /** Green: sets green channel (-100000 to 100000) */
  green?: number;
  /** Green offset: adjusts green channel (-100000 to 100000) */
  greenOff?: number;
  /** Green modulation: scales green channel (0-100000) */
  greenMod?: number;
  /** Blue: sets blue channel (-100000 to 100000) */
  blue?: number;
  /** Blue offset: adjusts blue channel (-100000 to 100000) */
  blueOff?: number;
  /** Blue modulation: scales blue channel (0-100000) */
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
 * createColorTransforms({ tint: 40000 });
 * // Semi-transparent red with 50% alpha
 * createColorTransforms({ alpha: 50000 });
 * ```
 */
export const createColorTransforms = (options: ColorTransformOptions): readonly string[] => {
  const t: string[] = [];

  if (options.tint !== undefined) t.push(`<a:tint val="${options.tint}"/>`);
  if (options.shade !== undefined) t.push(`<a:shade val="${options.shade}"/>`);
  if (options.comp) t.push(`<a:comp/>`);
  if (options.inv) t.push(`<a:inv/>`);
  if (options.gray) t.push(`<a:gray/>`);
  if (options.alpha !== undefined) t.push(`<a:alpha val="${options.alpha}"/>`);
  if (options.alphaOff !== undefined) t.push(`<a:alphaOff val="${options.alphaOff}"/>`);
  if (options.alphaMod !== undefined) t.push(`<a:alphaMod val="${options.alphaMod}"/>`);
  if (options.hue !== undefined) t.push(`<a:hue val="${options.hue}"/>`);
  if (options.hueOff !== undefined) t.push(`<a:hueOff val="${options.hueOff}"/>`);
  if (options.hueMod !== undefined) t.push(`<a:hueMod val="${options.hueMod}"/>`);
  if (options.sat !== undefined) t.push(`<a:sat val="${options.sat}"/>`);
  if (options.satOff !== undefined) t.push(`<a:satOff val="${options.satOff}"/>`);
  if (options.satMod !== undefined) t.push(`<a:satMod val="${options.satMod}"/>`);
  if (options.lum !== undefined) t.push(`<a:lum val="${options.lum}"/>`);
  if (options.lumOff !== undefined) t.push(`<a:lumOff val="${options.lumOff}"/>`);
  if (options.lumMod !== undefined) t.push(`<a:lumMod val="${options.lumMod}"/>`);
  if (options.red !== undefined) t.push(`<a:red val="${options.red}"/>`);
  if (options.redOff !== undefined) t.push(`<a:redOff val="${options.redOff}"/>`);
  if (options.redMod !== undefined) t.push(`<a:redMod val="${options.redMod}"/>`);
  if (options.green !== undefined) t.push(`<a:green val="${options.green}"/>`);
  if (options.greenOff !== undefined) t.push(`<a:greenOff val="${options.greenOff}"/>`);
  if (options.greenMod !== undefined) t.push(`<a:greenMod val="${options.greenMod}"/>`);
  if (options.blue !== undefined) t.push(`<a:blue val="${options.blue}"/>`);
  if (options.blueOff !== undefined) t.push(`<a:blueOff val="${options.blueOff}"/>`);
  if (options.blueMod !== undefined) t.push(`<a:blueMod val="${options.blueMod}"/>`);
  if (options.gamma) t.push(`<a:gamma/>`);
  if (options.invGamma) t.push(`<a:invGamma/>`);

  return t;
};
