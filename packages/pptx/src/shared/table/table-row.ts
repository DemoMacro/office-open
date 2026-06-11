import type { UniversalMeasure } from "@office-open/core";

import type { TableCellOptions } from "./table-cell";

export interface TableRowOptions {
  height?: number | UniversalMeasure;
  cells: TableCellOptions[];
}
