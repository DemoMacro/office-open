/**
 * Table cell module for WordprocessingML documents.
 *
 * Reference: http://officeopenxml.com/WPtableCell.php
 *
 * @module
 */

import type { SectionChild } from "@shared/section";

import type { TableCellPropertiesOptions } from "./table-cell-properties";

/**
 * Options for creating a TableCell element.
 *
 * @see {@link TableCell}
 */
export type TableCellOptions = {
  /** Array of Paragraph, Table, or plain objects that make up the cell content */
  children: SectionChild[];
} & TableCellPropertiesOptions;
