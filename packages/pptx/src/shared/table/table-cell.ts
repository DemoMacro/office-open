import type { BaseTableCellOptions, UniversalMeasure } from "@office-open/core";
import type { Cell3DOptions, ParagraphDescriptorOptions } from "@office-open/core/drawing";

import type { FillOptions } from "../drawing/fill";
import type { CellBorderOptions } from "./table-cell-properties";

/** Vertical alignment of cell content (ST_TextAnchorType); "distribute" spreads lines evenly across the cell height. */
export type VerticalAlignment = "top" | "center" | "bottom" | "justify" | "distribute";

/** ST_TextVerticalType — text direction within a cell (a:tcPr `@vert`). */
export type TextVerticalType = "horz" | "vert" | "vert270" | "wordArt" | "wordArtV";

/** pptx cell extends the base cell contract (span from base); verticalAlign
 *  widens to the DrawingML anchor set (justify/distribute) and fill/borders/
 *  margins/content are a:-domain types. */
export interface TableCellOptions extends Omit<BaseTableCellOptions, "verticalAlign"> {
  verticalAlign?: VerticalAlignment;
  /** `@vert` — text direction (ST_TextVerticalType). */
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
  /** Horizontal merge: "restart" first cell of the merge, "continue" cell absorbed into it. */
  horizontalMerge?: "continue" | "restart";
  /** Vertical merge: "restart" first cell of the merge, "continue" cell absorbed into it. */
  verticalMerge?: "continue" | "restart";
  /**
   * The source carried an `<a:tcPr/>` with no attributes or children.
   * Round-trip marker only — re-emits the empty element instead of dropping
   * it (real-world sources write a bare tcPr on every cell).
   */
  cellProperties?: boolean;
  margins?: {
    top?: number | UniversalMeasure;
    bottom?: number | UniversalMeasure;
    left?: number | UniversalMeasure;
    right?: number | UniversalMeasure;
  };
}
