/**
 * Tab stop module for WordprocessingML documents.
 *
 * This module provides tab stop definitions for paragraphs.
 *
 * Reference: http://officeopenxml.com/WPtab.php
 *
 * @module
 */

/**
 * Definition for a single tab stop.
 *
 * @see {@link TabStop}
 */
export interface TabStopDefinition {
  /** The type of tab stop alignment */
  type: (typeof TabStopType)[keyof typeof TabStopType];
  /** The position of the tab stop in twips */
  position: number | (typeof TabStopPosition)[keyof typeof TabStopPosition];
  /** Optional leader character to fill space before the tab */
  leader?: (typeof LeaderType)[keyof typeof LeaderType];
}

/**
 * Tab stop alignment types.
 *
 * Specifies the type of tab stop and how text aligns to it.
 *
 * @publicApi
 */
export const TabStopType = {
  /** Left-aligned tab stop */
  LEFT: "left",
  /** Right-aligned tab stop */
  RIGHT: "right",
  /** Center-aligned tab stop */
  CENTER: "center",
  /** Bar tab stop - inserts a vertical bar at the position */
  BAR: "bar",
  /** Clears a tab stop at the specified position */
  CLEAR: "clear",
  /** Decimal-aligned tab stop - aligns on decimal point */
  DECIMAL: "decimal",
  /** End-aligned tab stop (right-to-left equivalent) */
  END: "end",
  /** List tab stop for numbered lists */
  NUM: "num",
  /** Start-aligned tab stop (left-to-right equivalent) */
  START: "start",
} as const;

/**
 * Tab stop leader character types.
 *
 * Specifies the character used to fill the space before the tab stop.
 *
 * @publicApi
 */
export const LeaderType = {
  /** Dot leader (....) */
  DOT: "dot",
  /** Heavy leader */
  HEAVY: "heavy",
  /** Hyphen leader (----) */
  HYPHEN: "hyphen",
  /** Middle dot leader (····) */
  MIDDLE_DOT: "middleDot",
  /** No leader */
  NONE: "none",
  /** Underscore leader (____) */
  UNDERSCORE: "underscore",
} as const;

/**
 * Predefined tab stop positions.
 *
 * @publicApi
 */
export const TabStopPosition = {
  /** Maximum tab stop position (right margin) */
  MAX: 9026,
} as const;
