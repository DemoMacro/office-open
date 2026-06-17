/**
 * OOXML unit conversion utilities.
 *
 * @module
 */

// ---------------------------------------------------------------------------
// TWIP conversions (1 TWIP = 1/20 point, used in WordprocessingML)
// ---------------------------------------------------------------------------

/**
 * Converts millimeters to TWIP (twentieths of a point).
 */
export const convertMillimetersToTwip = (millimeters: number): number =>
  Math.floor((millimeters / 25.4) * 72 * 20);

/**
 * Converts inches to TWIP (twentieths of a point).
 */
export const convertInchesToTwip = (inches: number): number => Math.floor(inches * 72 * 20);

// ---------------------------------------------------------------------------
// EMU conversions (English Metric Units, used in DrawingML)
// 1 inch = 914400 EMU, 1 pixel (96 DPI) = 9525 EMU, 1 point = 12700 EMU
// ---------------------------------------------------------------------------

/**
 * Converts pixels to EMU (96 DPI).
 */
export const convertPixelsToEmu = (pixels: number): number => Math.round(pixels * 9525);

/**
 * Converts EMU to pixels (96 DPI).
 *
 * Returns a possibly fractional (sub-pixel) value. The integer rounding that
 * lived here before permanently discarded sub-pixel precision, which made an
 * EMU → pixel → EMU round-trip lossy (e.g. 5521960 EMU → 580 px → 5524500 EMU).
 * Keeping the fraction lets convertPixelsToEmu restore the exact original EMU.
 * Callers needing an integer pixel for display should Math.round the result.
 */
export const convertEmuToPixels = (emus: number): number => emus / 9525;

/**
 * Converts inches to EMU.
 */
export const convertInchesToEmu = (inches: number): number => Math.round(inches * 914400);

/**
 * Converts EMU to inches.
 */
export const convertEmuToInches = (emus: number): number => emus / 914400;

/**
 * Converts points to EMU.
 */
export const convertPointsToEmu = (points: number): number => Math.round(points * 12700);

/**
 * Converts EMU to points.
 */
export const convertEmuToPoints = (emus: number): number => emus / 12700;

// ---------------------------------------------------------------------------
// Position conversion (pixel-based position to EMU coordinates)
// ---------------------------------------------------------------------------

/** A rectangular position in pixels. */
export interface PixelPosition {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** An EMU-based rectangular position. */
export interface EmuPosition {
  x: number;
  y: number;
  cx: number;
  cy: number;
}

/**
 * Converts a pixel-based position to EMU coordinates.
 */
export const convertPositionToEmu = (pos: PixelPosition): EmuPosition => ({
  x: convertPixelsToEmu(pos.x),
  y: convertPixelsToEmu(pos.y),
  cx: convertPixelsToEmu(pos.width),
  cy: convertPixelsToEmu(pos.height),
});

// ---------------------------------------------------------------------------
// UniversalMeasure → Twips conversion
// Used when numeric computation is needed (e.g., landscape width/height swap)
// ---------------------------------------------------------------------------

/** Parsed result of a UniversalMeasure string. */
interface ParsedMeasure {
  value: number;
  unit: "mm" | "cm" | "in" | "pt" | "pc" | "pi";
}

/**
 * Parse a UniversalMeasure string into its numeric value and unit.
 *
 * @param measure - A universal measure string like "2.54cm", "-10mm", "1in"
 * @returns The parsed value and unit
 * @throws Error if the format is invalid
 *
 * @example
 * ```typescript
 * parseUniversalMeasure("2.54cm"); // { value: 2.54, unit: "cm" }
 * parseUniversalMeasure("-10mm");  // { value: -10, unit: "mm" }
 * ```
 */
export const parseUniversalMeasure = (measure: string): ParsedMeasure => {
  const match = measure.match(/^(-?[0-9]+(?:\.[0-9]+)?)(mm|cm|in|pt|pc|pi)$/);
  if (!match) {
    throw new Error(`Invalid universal measure: '${measure}'`);
  }
  return { value: parseFloat(match[1]), unit: match[2] as ParsedMeasure["unit"] };
};

/**
 * Converts a UniversalMeasure string to TWIP (twentieths of a point).
 *
 * Supports units: mm, cm, in, pt, pc (picas, 1pc = 12pt), pi (alias for pc).
 *
 * @param measure - A universal measure string like "2.54cm", "1in", "12pt"
 * @returns The value in TWIP
 *
 * @example
 * ```typescript
 * convertUniversalMeasureToTwip("1in");    // 1440
 * convertUniversalMeasureToTwip("2.54cm"); // ~1440 (1 inch)
 * convertUniversalMeasureToTwip("72pt");   // 1440 (1 inch = 72pt = 1440 twips)
 * ```
 */
export const convertUniversalMeasureToTwip = (measure: string): number => {
  const { value, unit } = parseUniversalMeasure(measure);
  switch (unit) {
    case "mm":
      return convertMillimetersToTwip(value);
    case "cm":
      return convertMillimetersToTwip(value * 10);
    case "in":
      return convertInchesToTwip(value);
    case "pt":
      return Math.floor(value * 20); // 1 point = 20 twips
    case "pc":
    case "pi":
      return Math.floor(value * 12 * 20); // 1 pica = 12 points = 240 twips
  }
};

/**
 * Converts a measurement value (number or UniversalMeasure) to TWIP.
 *
 * If the value is already a number, it is returned as-is (assumed to be in twips).
 * If the value is a UniversalMeasure string, it is converted to twips.
 *
 * Useful for accepting both `number` and `UniversalMeasure` inputs where
 * the XSD type is a union (e.g., ST_TwipsMeasure, ST_SignedTwipsMeasure).
 *
 * @param val - A numeric twip value or a universal measure string
 * @returns The value in TWIP
 *
 * @example
 * ```typescript
 * convertToTwip(1440);     // 1440 (already twips)
 * convertToTwip("1in");    // 1440
 * convertToTwip("2.54cm"); // ~1440
 * ```
 */
export const convertToTwip = (val: number | string): number =>
  typeof val === "string" ? convertUniversalMeasureToTwip(val) : val;

// ---------------------------------------------------------------------------
// UniversalMeasure → EMU conversion
// Used in DrawingML (PPTX shapes, images, etc.)
// ---------------------------------------------------------------------------

/**
 * Converts a UniversalMeasure string to EMU (English Metric Units).
 *
 * Supports units: mm, cm, in, pt, pc (picas, 1pc = 12pt), pi (alias for pc).
 *
 * @param measure - A universal measure string like "2.54cm", "1in", "12pt"
 * @returns The value in EMU
 *
 * @example
 * ```typescript
 * convertUniversalMeasureToEmu("1in");    // 914400
 * convertUniversalMeasureToEmu("2.54cm"); // 914400
 * convertUniversalMeasureToEmu("12pt");   // 152400
 * ```
 */
export const convertUniversalMeasureToEmu = (measure: string): number => {
  const { value, unit } = parseUniversalMeasure(measure);
  switch (unit) {
    case "mm":
      return Math.round(value * 36000);
    case "cm":
      return Math.round(value * 360000);
    case "in":
      return convertInchesToEmu(value);
    case "pt":
      return convertPointsToEmu(value);
    case "pc":
    case "pi":
      return convertPointsToEmu(value * 12); // 1 pica = 12 points
  }
};

/**
 * Converts a measurement value (number or UniversalMeasure) to EMU.
 *
 * If the value is already a number, it is returned as-is (assumed to be in EMU).
 * If the value is a UniversalMeasure string, it is converted to EMU.
 *
 * Useful for accepting both `number` and `UniversalMeasure` inputs in DrawingML
 * where the XSD type is a union (e.g., ST_Coordinate).
 *
 * @param val - A numeric EMU value or a universal measure string
 * @returns The value in EMU
 *
 * @example
 * ```typescript
 * convertToEmu(914400);   // 914400 (already EMU)
 * convertToEmu("1in");    // 914400
 * convertToEmu("2.54cm"); // 914400
 * ```
 */
export const convertToEmu = (val: number | string): number =>
  typeof val === "string" ? convertUniversalMeasureToEmu(val) : val;
