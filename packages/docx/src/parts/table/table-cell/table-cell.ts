/**
 * Table cell module for WordprocessingML documents.
 *
 * Reference: http://officeopenxml.com/WPtableCell.php
 *
 * @module
 */

import type { RunPropertiesOptions } from "@parts/paragraph/run/properties";
import type { SdtPropertiesOptions } from "@parts/table-of-contents";
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

/** Options for a cell-level structured document tag (CT_SdtCell). */
export interface SdtCellOptions {
  properties: SdtPropertiesOptions;
  /** Cells wrapped by the SDT (sdtContent holds <w:tc>). */
  cells?: TableCellOptions[];
  /** Run properties for the SDT end mark (w:sdtEndPr). */
  endProperties?: RunPropertiesOptions;
}
