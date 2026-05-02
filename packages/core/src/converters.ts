/**
 * OOXML unit conversion utilities.
 *
 * @module
 */

/**
 * Converts millimeters to TWIP (twentieths of a point).
 *
 * @publicApi
 */
export const convertMillimetersToTwip = (millimeters: number): number =>
    Math.floor((millimeters / 25.4) * 72 * 20);

/**
 * Converts inches to TWIP (twentieths of a point).
 *
 * @publicApi
 */
export const convertInchesToTwip = (inches: number): number => Math.floor(inches * 72 * 20);
