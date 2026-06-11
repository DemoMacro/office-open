/**
 * Conditional formatting style module for WordprocessingML documents.
 *
 * This module provides the conditional formatting style element (w:cnfStyle)
 * which specifies conditional formatting for table cells based on their position.
 *
 * Reference: http://officeopenxml.com/WPtableCellProperties.php
 *
 * @module
 */

/**
 * Options for conditional formatting style.
 *
 * These options specify which conditions apply to the element based on
 * its position within a table (first/last row, first/last column, bands).
 */
export interface CnfConditionalOptions {
  /** Whether this is the first row in the table */
  firstRow?: boolean;
  /** Whether this is the last row in the table */
  lastRow?: boolean;
  /** Whether this is the first column in the table */
  firstColumn?: boolean;
  /** Whether this is the last column in the table */
  lastColumn?: boolean;
  /** Whether this is an odd vertical band */
  oddVBand?: boolean;
  /** Whether this is an even vertical band */
  evenVBand?: boolean;
  /** Whether this is an odd horizontal band */
  oddHBand?: boolean;
  /** Whether this is an even horizontal band */
  evenHBand?: boolean;
  /** Whether this is the first row and first column (top-left corner) */
  firstRowFirstColumn?: boolean;
  /** Whether this is the first row and last column (top-right corner) */
  firstRowLastColumn?: boolean;
  /** Whether this is the last row and first column (bottom-left corner) */
  lastRowFirstColumn?: boolean;
  /** Whether this is the last row and last column (bottom-right corner) */
  lastRowLastColumn?: boolean;
}
