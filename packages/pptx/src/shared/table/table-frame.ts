import type {
  BaseTableOptions,
  EffectListOptions,
  FillOptions,
  GraphicFrameLockingOptions,
  NonVisualDrawingPropertiesOptions,
  TableStyleOptions,
  UniversalMeasure,
} from "@office-open/core";
import type { Guid } from "@office-open/core";
import type { GraphicFrameNvPrOptions } from "@parts/descriptors/graphic-frame";

import type { CellBorderOptions } from "./table-cell-properties";
import type { TableRowOptions } from "./table-row";

/**
 * Table (p:graphicFrame with a:tbl) for pptx slides. Structural base
 * (rows/columnWidths/6-flags) from BaseTableOptions; cNvPr fields from
 * NonVisualDrawingPropertiesOptions.
 */
export interface TableOptions
  extends
    BaseTableOptions<TableRowOptions>,
    NonVisualDrawingPropertiesOptions,
    GraphicFrameNvPrOptions {
  /** Table id (p:cNvPr `@id`). Auto-generated if omitted. */
  id?: number;
  /** Frame locking (a:graphicFrameLocks). undefined = fresh default
   * (noGrp="1"); null = empty cNvGraphicFramePr; object = explicit flags. */
  locking?: GraphicFrameLockingOptions | null;

  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  tableStyleId?: Guid;
  /** `@rtl` — table cell order is right-to-left. */
  rightToLeft?: boolean;
  /** Inline table style (a:tableStyle in a:tblPr) — alternative to tableStyleId. */
  tableStyle?: TableStyleOptions;
  /**
   * Office 2014 column stamps (a16:colId per a:gridCol extension list),
   * parallel to `BaseTableOptions.columnWidths` — round-trip.
   */
  columnIds?: string[];
  borders?: {
    top?: CellBorderOptions;
    bottom?: CellBorderOptions;
    left?: CellBorderOptions;
    right?: CellBorderOptions;
  };
  /** Table-level fill (a:tblPr/a:solidFill etc.) — round-trip. */
  fill?: FillOptions | null;
  /** Table-level effect list (a:tblPr/a:effectLst) — round-trip. */
  effects?: EffectListOptions;
}
