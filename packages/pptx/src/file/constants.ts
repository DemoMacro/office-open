/**
 * OOXML magic-number constants for PPTX generation.
 *
 * Values are in EMU (English Metric Units) or hundredths of a point,
 * as defined by the OOXML specification.
 * @module
 */

// ── Outline defaults (a:ln) ──

/** Default outline width: 1pt = 12700 EMU */
export const DEFAULT_OUTLINE_WIDTH = 12700;

// ── Shadow defaults (a:outerShdw) ──

/** Default shadow blur radius: ~4pt = 50800 EMU */
export const DEFAULT_SHADOW_BLUR_RADIUS = 50800;

/** Default shadow distance: ~3pt = 38100 EMU */
export const DEFAULT_SHADOW_DISTANCE = 38100;

/** Default shadow direction: 270° = 2700000 (60,000ths of a degree) */
export const DEFAULT_SHADOW_DIRECTION = 2700000;

/** Default shadow alpha: 40% = 40000 (100,000ths of a percent) */
export const DEFAULT_SHADOW_ALPHA = 40000;

// ── Angle / percentage multipliers ──

/** Degrees to OOXML angle units: 1° = 60,000 */
export const DEGREE_MULTIPLIER = 60000;

/** Font size to OOXML units: 1pt = 100 hundredths */
export const FONT_SIZE_MULTIPLIER = 100;

/** Percentage to OOXML units: 1% = 1,000 */
export const PERCENTAGE_MULTIPLIER = 1000;
