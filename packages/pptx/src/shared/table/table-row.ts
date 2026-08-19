import type { BaseTableRowOptions, UniversalMeasure } from "@office-open/core";

import type { TableCellOptions } from "./table-cell";

export interface TableRowOptions extends BaseTableRowOptions<TableCellOptions> {
  height?: number | UniversalMeasure;
  /** Office 2014 row stamp (a16:rowId in the a:tr extension list) — round-trip. */
  rowId?: string;
}
