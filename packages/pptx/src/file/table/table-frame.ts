import type { CellBorderOptions } from "./table-cell-properties";
import type { TableRowOptions } from "./table-row";

export interface TableOptions {
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly rows: readonly TableRowOptions[];
  readonly columnWidths?: readonly number[];
  readonly firstRow?: boolean;
  readonly lastRow?: boolean;
  readonly bandRow?: boolean;
  readonly firstCol?: boolean;
  readonly lastCol?: boolean;
  readonly bandCol?: boolean;
  readonly tableStyleId?: string;
  readonly borders?: {
    readonly top?: CellBorderOptions;
    readonly bottom?: CellBorderOptions;
    readonly left?: CellBorderOptions;
    readonly right?: CellBorderOptions;
  };
}
