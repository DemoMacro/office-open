import type { BaseTableRowOptions, UniversalMeasure } from "@office-open/core";

import type { TableCellOptions } from "./table-cell";

export interface TableRowOptions extends BaseTableRowOptions<TableCellOptions> {
  height?: number | UniversalMeasure;
}
