import type {
  BaseTableOptions,
  GraphicFrameLockingOptions,
  NonVisualDrawingPropertiesOptions,
  TableStyleOptions,
  UniversalMeasure,
} from "@office-open/core";
import type { NvPrPlaceholderOptions } from "@parts/descriptors/graphic-frame";

import type { CellBorderOptions } from "./table-cell-properties";
import type { TableRowOptions } from "./table-row";

/**
 * Table options for pptx slides (p:graphicFrame with a:tbl). The structural
 * base (rows/columnWidths/6-flags) comes from {@link BaseTableOptions}; the
 * cNvPr fields (name/description/title/hidden) from
 * {@link NonVisualDrawingPropertiesOptions}. The single source of truth for
 * both the public slide-child entry and the descriptor.
 */
export interface TableOptions
  extends
    BaseTableOptions<TableRowOptions>,
    NonVisualDrawingPropertiesOptions,
    NvPrPlaceholderOptions {
  /** Table id (p:cNvPr `@id`). Auto-generated if omitted. */
  id?: number;
  /** Frame locking (a:graphicFrameLocks). undefined = fresh default
   * (noGrp="1"); null = empty cNvGraphicFramePr; object = explicit flags. */
  locking?: GraphicFrameLockingOptions | null;

  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  tableStyleId?: string;
  /** Inline table style (a:tableStyle in a:tblPr) — alternative to tableStyleId. */
  tableStyle?: TableStyleOptions;
  borders?: {
    top?: CellBorderOptions;
    bottom?: CellBorderOptions;
    left?: CellBorderOptions;
    right?: CellBorderOptions;
  };
}
