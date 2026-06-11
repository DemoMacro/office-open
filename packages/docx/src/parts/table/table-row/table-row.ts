/**
 * Table row module for WordprocessingML documents.
 *
 * Reference: http://officeopenxml.com/WPtableRow.php
 *
 * @module
 */

import type { TableCellOptions } from "../table-cell";
import type { TablePropertyExOptions } from "../table-properties/table-property-exceptions";
import type { ITableRowPropertiesOptions } from "./table-row-properties";

/**
 * Options for creating a TableRow element.
 *
 * @see {@link TableRow}
 */
export type TableRowOptions = {
  /** Array of TableCell options that make up the row */
  cells: TableCellOptions[];
  /** Table property exceptions for this row (override table-level properties) */
  propertyExceptions?: TablePropertyExOptions;
  /** Revision save ID for row properties (hex string, e.g. "00123456"). */
  rsidRPr?: string;
  /** Revision save ID for the row (hex string). */
  rsidR?: string;
  /** Revision save ID when row was deleted (hex string). */
  rsidDel?: string;
  /** Revision save ID for table row (hex string). */
  rsidTr?: string;
} & ITableRowPropertiesOptions;
