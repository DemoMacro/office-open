import type { TableCellOptions } from "./table-cell";

export interface TableRowOptions {
  readonly height?: number;
  readonly cells: readonly TableCellOptions[];
}
