import type { FillOptions } from "../drawingml/fill";
import type { ParagraphOptions } from "../shape/paragraph/paragraph";
import type { CellBorderOptions } from "./table-cell-properties";

export type VerticalAlignment = "top" | "center" | "bottom" | "justify" | "distribute";

export interface TableCellOptions {
  readonly text?: string;
  readonly children?: readonly (ParagraphOptions | string)[];
  readonly fill?: FillOptions;
  readonly borders?: {
    readonly top?: CellBorderOptions;
    readonly bottom?: CellBorderOptions;
    readonly left?: CellBorderOptions;
    readonly right?: CellBorderOptions;
  };
  readonly columnSpan?: number;
  readonly rowSpan?: number;
  readonly horizontalMerge?: "continue" | "restart";
  readonly verticalMerge?: "continue" | "restart";
  readonly verticalAlign?: VerticalAlignment;
  readonly margins?: {
    readonly top?: number;
    readonly bottom?: number;
    readonly left?: number;
    readonly right?: number;
  };
}
