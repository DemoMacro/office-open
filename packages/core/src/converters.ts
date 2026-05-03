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
 */
export const convertEmuToPixels = (emus: number): number => Math.round(emus / 9525);

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
