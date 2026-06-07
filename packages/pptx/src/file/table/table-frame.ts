import type { CellBorderOptions } from "./table-cell-properties";
import type { TableRowOptions } from "./table-row";

export interface TableOptions {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  rows: TableRowOptions[];
  columnWidths?: number[];
  firstRow?: boolean;
  lastRow?: boolean;
  bandRow?: boolean;
  firstCol?: boolean;
  lastCol?: boolean;
  bandCol?: boolean;
  tableStyleId?: string;
  borders?: {
    top?: CellBorderOptions;
    bottom?: CellBorderOptions;
    left?: CellBorderOptions;
    right?: CellBorderOptions;
  };
}
