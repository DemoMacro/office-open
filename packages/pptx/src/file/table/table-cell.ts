import type { FillOptions } from "../drawingml/fill";
import type { ParagraphOptions } from "../shape/paragraph/paragraph";
import type { CellBorderOptions } from "./table-cell-properties";

export type VerticalAlignment = "top" | "center" | "bottom" | "justify" | "distribute";

export interface TableCellOptions {
  text?: string;
  children?: (ParagraphOptions | string)[];
  fill?: FillOptions;
  borders?: {
    top?: CellBorderOptions;
    bottom?: CellBorderOptions;
    left?: CellBorderOptions;
    right?: CellBorderOptions;
  };
  columnSpan?: number;
  rowSpan?: number;
  horizontalMerge?: "continue" | "restart";
  verticalMerge?: "continue" | "restart";
  verticalAlign?: VerticalAlignment;
  margins?: {
    top?: number;
    bottom?: number;
    left?: number;
    right?: number;
  };
}
