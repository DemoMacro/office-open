/**
 * Table module for WordprocessingML documents.
 *
 * Reference: http://officeopenxml.com/WPtableGrid.php
 *
 * @module
 */

import type { PositiveUniversalMeasure } from "@office-open/core";

import type { AlignmentType } from "../paragraph";
import type { TableGridChangeOptions } from "./grid";
import type { TableCellSpacingProperties } from "./table-cell-spacing";
import type {
  TableBordersOptions,
  TableFloatOptions,
  TablePropertiesChangeOptions,
} from "./table-properties";
import type { TableCellMarginOptions } from "./table-properties/table-cell-margin";
import type { TableLayoutType } from "./table-properties/table-layout";
import type { TableLookOptions } from "./table-properties/table-look";
import type { TableRowOptions } from "./table-row";
import type { TableWidthProperties } from "./table-width";

/**
 * Options for creating a Table element.
 *
 * Note: 0-width columns don't get rendered correctly, so we need
 * to give them some value. A reasonable default would be
 * ~6in / numCols, but if we do that it becomes very hard
 * to resize the table using setWidth, unless the layout
 * algorithm is set to 'fixed'. Instead, the approach here
 * means even in 'auto' layout, setting a width on the
 * table will make it look reasonable, as the layout
 * algorithm will expand columns to fit its content.
 *
 * @see {@link Table}
 */
export interface TableOptions {
  rows: TableRowOptions[];
  width?: TableWidthProperties;
  columnWidths?: number[] | PositiveUniversalMeasure[];
  columnWidthsRevision?: TableGridChangeOptions;
  margins?: TableCellMarginOptions;
  indent?: TableWidthProperties;
  float?: TableFloatOptions;
  layout?: (typeof TableLayoutType)[keyof typeof TableLayoutType];
  style?: string;
  borders?: TableBordersOptions;
  alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
  visuallyRightToLeft?: boolean;
  tableLook?: TableLookOptions;
  cellSpacing?: TableCellSpacingProperties;
  styleRowBandSize?: number;
  styleColBandSize?: number;
  caption?: string;
  description?: string;
  revision?: TablePropertiesChangeOptions;
}
