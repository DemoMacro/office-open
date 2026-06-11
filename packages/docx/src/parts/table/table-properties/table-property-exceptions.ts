/**
 * Table property exceptions module for WordprocessingML documents.
 *
 * Property exceptions allow overriding table properties at the row level.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_TblPrExBase
 *
 * @module
 */

import type { ShadingAttributesProperties } from "@shared/shading";

import type { AlignmentType } from "../../paragraph";
import type { TableCellSpacingProperties } from "../table-cell-spacing";
import type { TableWidthProperties } from "../table-width";
import type { TableBordersOptions } from "./table-borders";
import type { TableCellMarginOptions } from "./table-cell-margin";
import type { TableLayoutType } from "./table-layout";
import type { TableLookOptions } from "./table-look";

/**
 * Options for table property exceptions (w:tblPrEx).
 *
 * These override the parent table's properties for a specific row.
 * Subset of TablePropertiesOptionsBase.
 */
export interface TablePropertyExOptions {
  width?: TableWidthProperties;
  indent?: TableWidthProperties;
  layout?: (typeof TableLayoutType)[keyof typeof TableLayoutType];
  borders?: TableBordersOptions;
  shading?: ShadingAttributesProperties;
  alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
  cellMargin?: TableCellMarginOptions;
  tableLook?: TableLookOptions;
  cellSpacing?: TableCellSpacingProperties;
  /** Table property exceptions change tracking */
  tblPrExChange?: { author: string; date?: string; id: string };
}
