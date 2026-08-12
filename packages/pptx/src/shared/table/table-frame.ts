import type { BaseTableOptions, UniversalMeasure } from "@office-open/core";

import type { CellBorderOptions } from "./table-cell-properties";
import type { TableRowOptions } from "./table-row";

export interface TableOptions extends BaseTableOptions<TableRowOptions> {
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
