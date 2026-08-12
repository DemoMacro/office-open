import type { BaseTableCellOptions, UniversalMeasure } from "@office-open/core";
import type { ParagraphDescriptorOptions } from "@office-open/core/drawingml";

import type { FillOptions } from "../drawingml/fill";
import type { CellBorderOptions } from "./table-cell-properties";

export type VerticalAlignment = "top" | "center" | "bottom" | "justify" | "distribute";

/** pptx cell extends the base cell contract (span from base); verticalAlign
 *  widens to the DrawingML anchor set (justify/distribute) and fill/borders/
 *  margins/content are a:-domain types. */
export interface TableCellOptions extends Omit<BaseTableCellOptions, "verticalAlign"> {
  verticalAlign?: VerticalAlignment;
  text?: string;
  children?: (ParagraphDescriptorOptions | string)[];
  fill?: FillOptions;
  borders?: {
    top?: CellBorderOptions;
    bottom?: CellBorderOptions;
    left?: CellBorderOptions;
    right?: CellBorderOptions;
    diagonalTopLeftToBottomRight?: CellBorderOptions;
    diagonalBottomLeftToTopRight?: CellBorderOptions;
  };
  horizontalMerge?: "continue" | "restart";
  verticalMerge?: "continue" | "restart";
  margins?: {
    top?: number | UniversalMeasure;
    bottom?: number | UniversalMeasure;
    left?: number | UniversalMeasure;
    right?: number | UniversalMeasure;
  };
}
