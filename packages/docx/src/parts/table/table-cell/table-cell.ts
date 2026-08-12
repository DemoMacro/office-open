/**
 * Table cell module for WordprocessingML documents.
 *
 * Reference: http://officeopenxml.com/WPtableCell.php
 *
 * @module
 */

import type { BaseTableCellOptions } from "@office-open/core";
import type { RunPropertiesOptions } from "@parts/paragraph/run/properties";
import type { SdtPropertiesOptions } from "@parts/table-of-contents";
import type { SectionChild } from "@shared/section";

import type { TableCellPropertiesOptions } from "./table-cell-properties";

/**
 * Options for creating a TableCell element.
 *
 * Extends {@link BaseTableCellOptions} for the cross-format cell contract
 * (span + vertical-align); shading/borders/margins/content stay docx-specific.
 *
 * @see {@link TableCell}
 */
export interface TableCellOptions extends BaseTableCellOptions, TableCellPropertiesOptions {
  /** Array of Paragraph, Table, or plain objects that make up the cell content */
  children: SectionChild[];
}

/** Options for a cell-level structured document tag (CT_SdtCell). */
export interface SdtCellOptions {
  properties: SdtPropertiesOptions;
  /** Cells wrapped by the SDT (sdtContent holds <w:tc>). */
  cells?: TableCellOptions[];
  /** Run properties for the SDT end mark (w:sdtEndPr). */
  endProperties?: RunPropertiesOptions;
}
