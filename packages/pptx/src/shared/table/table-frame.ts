import type {
  BaseTableOptions,
  NonVisualDrawingPropertiesOptions,
  UniversalMeasure,
} from "@office-open/core";

import type { CellBorderOptions } from "./table-cell-properties";
import type { TableRowOptions } from "./table-row";

/**
 * Table options for pptx slides (p:graphicFrame with a:tbl). The structural
 * base (rows/columnWidths/6-flags) comes from {@link BaseTableOptions}; the
 * cNvPr fields (name/description/title/hidden) from
 * {@link NonVisualDrawingPropertiesOptions}. The single source of truth for
 * both the public slide-child entry and the descriptor.
 */
export interface TableOptions
  extends BaseTableOptions<TableRowOptions>, NonVisualDrawingPropertiesOptions {
  /** Table id (p:cNvPr @id). Auto-generated if omitted. */
  id?: number;
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  tableStyleId?: string;
  borders?: {
    top?: CellBorderOptions;
    bottom?: CellBorderOptions;
    left?: CellBorderOptions;
    right?: CellBorderOptions;
  };
}
