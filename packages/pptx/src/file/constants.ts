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

// ── Shape tree defaults (p:spTree) ──

/** Empty shape tree header: nvGrpSpPr + grpSpPr with zero-offset transform */
export const SP_TREE_HEADER =
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
  '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

// ── Default color map (p:clrMap) ──

/** Default theme color mapping attribute string for <p:clrMap> */
export const DEFAULT_COLOR_MAP =
  'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"';
