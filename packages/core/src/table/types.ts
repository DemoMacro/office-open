/**
 * Base table options shared across docx (w:tbl) and pptx (a:tbl).
 *
 * Only structurally-compatible fields live here. Style fields (fill/borders/
 * margins), row height, and cell content stay per-package: docx uses w:-domain
 * types (ShadingProperties, TableCellBordersOptions, TableCellMarginOptions,
 * {value,rule} height, SectionChild content); pptx uses a:-domain types
 * (FillOptions, CellBorderOptions, 4-side margins, number|UM height,
 * ParagraphDescriptorOptions content). The two XML namespaces cannot be
 * unified without loss, so each package extends this base and adds its own.
 *
 * The row/cell collections are generic so each package can carry its own
 * element union — docx wraps rows and cells in sdt/customXml containers,
 * pptx uses plain row/cell arrays.
 *
 * @module
 */
import type { PositiveUniversalMeasure } from "../util/values";

/** Cell fields shared by docx and pptx (span + vertical-align intersection). */
export interface BaseTableCellOptions {
  /** Columns this cell spans (docx gridSpan/hMerge, pptx gridSpan/hMerge). */
  columnSpan?: number;
  /** Rows this cell spans (docx vMerge restart, pptx rowSpan/vMerge). */
  rowSpan?: number;
  /** Vertical alignment — docx/pptx intersection. pptx adds justify/distribute. */
  verticalAlign?: "top" | "center" | "bottom";
}

/** Row fields shared by docx and pptx. Generic in cell type so each package
 *  can carry its own cell union (docx adds sdt/customXml wrappers). */
export interface BaseTableRowOptions<TCell = BaseTableCellOptions> {
  cells: TCell[];
}

/** Table fields shared by docx and pptx. Generic in row type so each package
 *  can carry its own row union (docx adds sdt/customXml wrappers). */
export interface BaseTableOptions<TRow = BaseTableRowOptions> {
  rows: TRow[];
  /**
   * Column widths. `number` is the package's native unit (docx twip, pptx EMU);
   * `PositiveUniversalMeasure` is consistent across packages. Cross-format
   * conversion translates the native unit.
   */
  columnWidths?: (number | PositiveUniversalMeasure)[];
  /** Apply first-row conditional formatting (docx w:tblLook@firstRow, pptx a:tblPr@firstRow). */
  firstRow?: boolean;
  lastRow?: boolean;
  firstCol?: boolean;
  lastCol?: boolean;
  bandRow?: boolean;
  bandCol?: boolean;
}
