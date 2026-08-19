/**
 * Table row module for WordprocessingML documents.
 *
 * Reference: http://officeopenxml.com/WPtableRow.php
 *
 * @module
 */

import type { BaseTableRowOptions } from "@office-open/core";
import type { CustomXmlCellOptions } from "@parts/custom-xml";
import type { BookmarkStartOptions, MarkupRangeOptions } from "@parts/paragraph/links/bookmark";
import type { RunPropertiesOptions } from "@parts/paragraph/run/properties";
import type { SdtPropertiesOptions } from "@parts/table-of-contents";

import type { SdtCellOptions, TableCellOptions } from "../table-cell";
import type { TablePropertyExOptions } from "../table-properties/table-property-exceptions";
import type { TableRowPropertiesOptions } from "./table-row-properties";

/**
 * Run-level markers anchored directly under `w:tr`, between a row's cells —
 * comment and bookmark ranges spanning cell boundaries land there because
 * the tc content model cannot hold them at that position.
 */
export type TableRowMarkerChild =
  | { commentRangeStart: MarkupRangeOptions }
  | { commentRangeEnd: MarkupRangeOptions }
  | { bookmarkStart: BookmarkStartOptions }
  | { bookmarkEnd: MarkupRangeOptions };

/**
 * Options for creating a TableRow element.
 *
 * @see {@link TableRow}
 */
/** Options for a row-level structured document tag (CT_SdtRow). */
export interface SdtRowOptions {
  properties: SdtPropertiesOptions;
  /** Rows wrapped by the SDT (sdtContent holds <w:tr>). */
  rows?: TableRowOptions[];
  /** Run properties for the SDT end mark (w:sdtEndPr). */
  endProperties?: RunPropertiesOptions;
}

export interface TableRowOptions
  extends
    BaseTableRowOptions<
      | TableCellOptions
      | { sdt: SdtCellOptions }
      | { customXml: CustomXmlCellOptions }
      | TableRowMarkerChild
    >,
    TableRowPropertiesOptions {
  /** Table property exceptions for this row (override table-level properties) */
  propertyExceptions?: TablePropertyExOptions;
  /** Revision save ID for row properties (w:rsidRPr, hex string). */
  runPropertiesRsid?: string;
  /** Revision save ID for the row (w:rsidR, hex string). */
  additionRsid?: string;
  /** Revision save ID when row was deleted (w:rsidDel, hex string). */
  deletionRsid?: string;
  /** Revision save ID for table row (w:rsidTr, hex string). */
  tableRowRsid?: string;
}
