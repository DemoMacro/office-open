/**
 * Text wrapping module for DrawingML elements.
 *
 * This module provides text wrapping options for floating/anchored drawings.
 *
 * Reference: http://officeopenxml.com/drwPicFloating-textWrap.php
 *
 * @module
 */
import type { Distance } from "../drawing";

/**
 * Enumeration of text wrapping types for floating drawings.
 *
 * Reference: http://officeopenxml.com/drwPicFloating-textWrap.php
 *
 * @publicApi
 */
export const TextWrappingType = {
  NONE: 0,
  SQUARE: 1,
  TIGHT: 2,
  TOP_AND_BOTTOM: 3,
  THROUGH: 4,
} as const;

/**
 * Enumeration of text wrapping sides for floating drawings.
 *
 * Specifies on which side(s) text can wrap around the drawing.
 *
 * Reference: http://officeopenxml.com/drwPicFloating-textWrap.php
 *
 * @publicApi
 */
export const TextWrappingSide = {
  /** Text wraps on both sides of the drawing */
  BOTH_SIDES: "bothSides",
  /** Text wraps only on the left side */
  LEFT: "left",
  /** Text wraps only on the right side */
  RIGHT: "right",
  /** Text wraps on the side with more space */
  LARGEST: "largest",
} as const;

/**
 * A point in a wrap polygon (wrapTight/wrapThrough), in EMUs.
 */
export interface WrapPolygonPoint {
  x: number;
  y: number;
}

/**
 * Wrap polygon for wrapTight/wrapThrough — the contour text wraps around.
 * Round-tripped verbatim so Word's authored contour is preserved.
 */
export interface WrapPolygon {
  edited?: boolean;
  points: WrapPolygonPoint[];
}

/**
 * Options for configuring text wrapping around a drawing.
 */
export interface TextWrapping {
  type: (typeof TextWrappingType)[keyof typeof TextWrappingType];
  side?: (typeof TextWrappingSide)[keyof typeof TextWrappingSide];
  margins?: Distance;
  /** Wrap polygon for wrapTight/wrapThrough. Preserves the source contour on round-trip; defaults to the extent rectangle when unset. */
  polygon?: WrapPolygon;
}
