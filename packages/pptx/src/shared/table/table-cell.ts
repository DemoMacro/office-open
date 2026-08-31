import type { BaseTableCellOptions, UniversalMeasure } from "@office-open/core";
import type {
  Cell3DOptions,
  ParagraphDescriptorOptions,
  VerticalAnchor,
} from "@office-open/core/drawing";

import type { FillOptions } from "../drawing/fill";
import type { CellBorderOptions } from "./table-cell-properties";

// Vertical anchor of cell content is the core DrawingML token set
// (ST_TextAnchorType); re-exported under the table-cell domain.
export type { VerticalAnchor } from "@office-open/core/drawing";

/** Text direction within a cell (ST_TextVerticalType, a:tcPr `@vert`). */
export type TextVerticalType =
  | "horizontal"
  | "vertical"
  | "vertical270"
  | "wordArtVertical"
  | "wordArtVerticalRightToLeft";

/** pptx cell extends the base cell contract (span from base); verticalAlign
 *  widens to the DrawingML anchor set (justify/distribute) and fill/borders/
 *  margins/content are a:-domain types. */
export interface TableCellOptions extends Omit<BaseTableCellOptions, "verticalAlign"> {
  verticalAlign?: VerticalAnchor;
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
