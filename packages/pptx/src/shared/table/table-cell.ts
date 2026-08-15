import type { BaseTableCellOptions, UniversalMeasure } from "@office-open/core";
import type { Cell3DOptions, ParagraphDescriptorOptions } from "@office-open/core/drawing";

import type { FillOptions } from "../drawing/fill";
import type { CellBorderOptions } from "./table-cell-properties";

export type VerticalAlignment = "top" | "center" | "bottom" | "justify" | "distribute";

/** ST_TextVerticalType — text direction within a cell (a:tcPr @vert). */
export type TextVerticalType = "horz" | "vert" | "vert270" | "wordArt" | "wordArtV";

/** pptx cell extends the base cell contract (span from base); verticalAlign
 *  widens to the DrawingML anchor set (justify/distribute) and fill/borders/
 *  margins/content are a:-domain types. */
export interface TableCellOptions extends Omit<BaseTableCellOptions, "verticalAlign"> {
  verticalAlign?: VerticalAlignment;
  /** @vert — text direction (ST_TextVerticalType). */
  vertical?: TextVerticalType;
  text?: string;
  children?: (ParagraphDescriptorOptions | string)[];
  fill?: FillOptions;
  /** Cell bevel (a:cell3D in a:tcPr). */
  cell3D?: Cell3DOptions;
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
