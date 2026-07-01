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
 * w:tblCellSpacing is a CT_TblWidth (same type as w:tblW/w:tcW/w:tblInd); its
 * @w:type is ST_TblWidth (nil/pct/dxa/auto). Cell spacing is conventionally dxa,
 * but pct is supported for parity with the other CT_TblWidth fields.
 */
export const CellSpacingType = {
  /** Value is in twentieths of a point (dxa) */
  DXA: "dxa",
  /** No (empty) value. */
  NIL: "nil",
  /** Value is a percentage of the table width (100 = 100%). */
  PERCENTAGE: "pct",
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
