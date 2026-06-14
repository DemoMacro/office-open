/**
 * Table cell spacing module for WordprocessingML documents.
 *
 * This module provides cell spacing settings for tables, controlling
 * the space between cells in a table.
 *
 * Reference: http://officeopenxml.com/WPtableCellSpacing.php
 *
 * @module
 */
import type { Percentage, UniversalMeasure } from "@office-open/core";

/**
 * Cell spacing measurement types.
 *
 * w:tblCellSpacing is a CT_TblWidth whose @w:type is ST_TblWidthType; cell
 * spacing only ever uses the nil/dxa subset of that union.
 */
export const CellSpacingType = {
  /** Value is in twentieths of a point (dxa) */
  DXA: "dxa",
  /** No (empty) value. */
  NIL: "nil",
} as const;

/**
 * Properties for table cell spacing.
 */
export interface TableCellSpacingProperties {
  /** The spacing value (in twips, percentage, or universal measure) */
  size: number | Percentage | UniversalMeasure;
  /** The type of measurement (defaults to DXA/twips) */
  type?: (typeof CellSpacingType)[keyof typeof CellSpacingType];
}
