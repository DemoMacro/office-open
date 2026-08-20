/**
 * Table borders module for WordprocessingML documents.
 *
 * This module provides border options for tables.
 *
 * Reference: http://officeopenxml.com/WPtableBorders.php
 *
 * @module
 */
import { BorderStyle } from "@shared/border";
import type { BorderOptions } from "@shared/border";

/**
 * Options for configuring table borders.
 *
 * Borders can be applied to the outside edges (top, bottom, left, right,
 * start, end) and inside lines (insideHorizontal, insideVertical) of the
 * table. start/end are the bidi-aware equivalents of left/right (CT_TblBorders
 * carries both; Word emits whichever the table's direction selects).
 */
export interface TableBordersOptions {
  top?: BorderOptions;
  bottom?: BorderOptions;
  left?: BorderOptions;
  right?: BorderOptions;
  start?: BorderOptions;
  end?: BorderOptions;
  insideHorizontal?: BorderOptions;
  insideVertical?: BorderOptions;
}

const NONE_BORDER: BorderOptions = {
  color: "auto",
  size: 0,
  style: BorderStyle.NONE,
};

export const TABLE_BORDERS_NONE: TableBordersOptions = {
  bottom: NONE_BORDER,
  insideHorizontal: NONE_BORDER,
  insideVertical: NONE_BORDER,
  left: NONE_BORDER,
  right: NONE_BORDER,
  top: NONE_BORDER,
};
