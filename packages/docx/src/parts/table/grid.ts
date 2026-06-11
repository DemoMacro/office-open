/**
 * Table grid module for WordprocessingML documents.
 *
 * The table grid defines the column structure of a table.
 *
 * Reference: http://officeopenxml.com/WPtableGrid.php
 *
 * @module
 */
import type { PositiveUniversalMeasure } from "@office-open/core";

export interface TableGridChangeOptions {
  id: number;
  columnWidths: number[] | PositiveUniversalMeasure[];
}
