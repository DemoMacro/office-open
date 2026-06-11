/**
 * Table cell components module for WordprocessingML documents.
 *
 * This module provides types and constants for table cell properties including borders,
 * grid span (column span), vertical merge, and text direction.
 *
 * Reference: http://officeopenxml.com/WPtableCell.php
 *
 * @module
 */
import type { BorderOptions } from "@shared/border";

/**
 * Options for configuring table cell borders.
 *
 * Defines border settings for individual edges of a table cell.
 */
export interface TableCellBordersOptions {
  /** Border for the top edge of the cell */
  top?: BorderOptions;
  /** Border for the start edge (left in LTR, right in RTL) */
  start?: BorderOptions;
  /** Border for the left edge of the cell */
  left?: BorderOptions;
  /** Border for the bottom edge of the cell */
  bottom?: BorderOptions;
  /** Border for the end edge (right in LTR, left in RTL) */
  end?: BorderOptions;
  /** Border for the right edge of the cell */
  right?: BorderOptions;
  /** Diagonal border from top-left to bottom-right */
  topLeftToBottomRight?: BorderOptions;
  /** Diagonal border from top-right to bottom-left */
  topRightToBottomLeft?: BorderOptions;
}

/**
 * Vertical merge types for table cells.
 *
 * Defines the merge behavior for vertically merged cells (row span).
 */
export const VerticalMergeType = {
  /**
   * Cell that is merged with upper one.
   */
  CONTINUE: "continue",
  /**
   * Cell that is starting the vertical merge.
   */
  RESTART: "restart",
} as const;

/**
 * Text direction values for table cells.
 *
 * Specifies the direction in which text flows within a table cell.
 */
export const TextDirection = {
  /** Text flows from bottom to top, left to right */
  BOTTOM_TO_TOP_LEFT_TO_RIGHT: "btLr",
  /** Text flows from left to right, top to bottom (default) */
  LEFT_TO_RIGHT_TOP_TO_BOTTOM: "lrTb",
  /** Text flows from top to bottom, right to left */
  TOP_TO_BOTTOM_RIGHT_TO_LEFT: "tbRl",
} as const;
