import type { ShadingAttributesProperties } from "@shared/shading";
/**
 * Table cell properties module for WordprocessingML documents.
 *
 * This module provides cell-level properties including width, borders,
 * shading, margins, and merge settings.
 *
 * Reference: http://officeopenxml.com/WPtableCellProperties.php
 *
 * @module
 */
import type { ICellMergeAttributes } from "@shared/track-revision";
import type { ChangedAttributesProperties } from "@shared/track-revision/track-revision";
import type { TableVerticalAlign } from "@shared/vertical-align";

import type { TableCellMarginOptions } from "../table-properties/table-cell-margin";
import type { CnfStyleOptions } from "../table-row/table-row-properties";
import type { TableWidthProperties } from "../table-width";
import type { TableCellBordersOptions, TextDirection } from "./table-cell-components";

export interface TableCellPropertiesOptionsBase {
  /** Conditional formatting style (cnfStyle) */
  cnfStyle?: CnfStyleOptions;
  /** Shading (background color/pattern) for the cell */
  shading?: ShadingAttributesProperties;
  /** Cell margins (padding) for the cell content */
  margins?: TableCellMarginOptions;
  /** Vertical alignment of content within the cell */
  verticalAlign?: TableVerticalAlign;
  /** Text direction/flow within the cell */
  textDirection?: (typeof TextDirection)[keyof typeof TextDirection];
  /** Vertical merge setting for the cell */
  verticalMerge?: "continue" | "restart";
  /** Width specification for the cell */
  width?: TableWidthProperties;
  /** Number of columns this cell spans (horizontal merge) */
  columnSpan?: number;
  /** Number of rows this cell spans (vertical merge) */
  rowSpan?: number;
  /** Border settings for the cell edges */
  borders?: TableCellBordersOptions;
  /** Horizontal merge setting (hMerge) */
  horizontalMerge?: "continue" | "restart";
  /** Whether the cell content does not wrap (noWrap) */
  noWrap?: boolean;
  /** Whether text is auto-fit to cell width (tcFitText) */
  fitText?: boolean;
  /** Whether the cell end mark is hidden (hideMark) */
  hideMark?: boolean;
  /** Header cells associated with this cell (headers) */
  headers?: string[];
  insertion?: ChangedAttributesProperties;
  deletion?: ChangedAttributesProperties;
  cellMerge?: ICellMergeAttributes;
}

/**
 * Options for configuring table cell properties.
 *
 * @see {@link TableCellProperties}
 */
export type ITableCellPropertiesOptions = {
  revision?: ITableCellPropertiesChangeOptions;
  includeIfEmpty?: boolean;
} & TableCellPropertiesOptionsBase;

export type ITableCellPropertiesChangeOptions = TableCellPropertiesOptionsBase &
  ChangedAttributesProperties;
