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
// UniversalMeasure → Twips conversion
// Used when numeric computation is needed (e.g., landscape width/height swap)
// ---------------------------------------------------------------------------

/** Parsed result of a UniversalMeasure string. */
interface ParsedMeasure {
  value: number;
  unit: "mm" | "cm" | "in" | "pt" | "pc" | "pi" | "px";
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
  const match = measure.match(/^(-?[0-9]+(?:\.[0-9]+)?)(mm|cm|in|pt|pc|pi|px)$/);
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
    case "px":
      return Math.round(value * 15); // 1px = 15twip (1in = 1440twip = 96px)
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
    case "px":
      return convertPixelsToEmu(value); // 1px = 9525 EMU (96 DPI)
  }
};

/**
 * Converts a measurement value (number or UniversalMeasure) to EMU.
 *
 * Numbers are returned as-is (assumed EMU). Strings are parsed as
 * UniversalMeasure via {@link convertUniversalMeasureToEmu} — including the
 * project-only `px` unit (96 DPI). The result is always an EMU number written to
 * XML, so px never appears verbatim in the document.
 *
 * Useful for DrawingML fields where the XSD type is a union (e.g., ST_Coordinate).
 *
 * @param val - A numeric EMU value, or a UniversalMeasure string (incl. `${n}px`)
 * @returns The value in EMU
 *
 * @example
 * ```typescript
 * convertToEmu(914400);    // 914400 (already EMU)
 * convertToEmu("1in");     // 914400
 * convertToEmu("2.54cm");  // 914400
 * convertToEmu("200px");   // 1905000 (200 * 9525)
 * ```
 */
export const convertToEmu = (val: number | string): number =>
  typeof val === "string" ? convertUniversalMeasureToEmu(val) : val;

// ---------------------------------------------------------------------------
// UniversalMeasure → Points / Inches conversion
// Used in SpreadsheetML (row height in points, page margins in inches)
// ---------------------------------------------------------------------------

/**
 * Converts a UniversalMeasure string to points (1pt = 1/72 inch).
 *
 * Supports units: mm, cm, in, pt, pc (picas, 1pc = 12pt), pi (alias for pc),
 * px (96 DPI).
 */
export const convertUniversalMeasureToPt = (measure: string): number => {
  const { value, unit } = parseUniversalMeasure(measure);
  switch (unit) {
    case "mm":
      return (value / 25.4) * 72;
    case "cm":
      return ((value * 10) / 25.4) * 72;
    case "in":
      return value * 72;
    case "pt":
      return value;
    case "pc":
    case "pi":
      return value * 12; // 1 pica = 12 points
    case "px":
      return (value / 96) * 72; // 1px = 0.75pt (96 DPI)
  }
};

/**
 * Converts a measurement value (number or UniversalMeasure) to points.
 *
 * Numbers are returned as-is (assumed to be in points). Strings are parsed as
 * UniversalMeasure. Useful for SpreadsheetML fields where a number is points.
 */
export const convertToPt = (val: number | string): number =>
  typeof val === "string" ? convertUniversalMeasureToPt(val) : val;

/**
 * Converts a UniversalMeasure string to inches.
 *
 * Supports units: mm, cm, in, pt, pc, pi, px (96 DPI).
 */
export const convertUniversalMeasureToInch = (measure: string): number => {
  const { value, unit } = parseUniversalMeasure(measure);
  switch (unit) {
    case "mm":
      return value / 25.4;
    case "cm":
      return (value * 10) / 25.4;
    case "in":
      return value;
    case "pt":
      return value / 72;
    case "pc":
    case "pi":
      return (value * 12) / 72;
    case "px":
      return value / 96;
  }
};

/**
 * Converts a measurement value (number or UniversalMeasure) to inches.
 *
 * Numbers are returned as-is (assumed to be in inches). Useful for SpreadsheetML
 * page-margin fields where a number is inches.
 */
export const convertToInch = (val: number | string): number =>
  typeof val === "string" ? convertUniversalMeasureToInch(val) : val;
