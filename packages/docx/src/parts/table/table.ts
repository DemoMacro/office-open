/**
 * Table module for WordprocessingML documents.
 *
 * Reference: http://officeopenxml.com/WPtableGrid.php
 *
 * @module
 */

import type { BaseTableOptions, PositiveUniversalMeasure } from "@office-open/core";
import type { CustomXmlRowOptions } from "@parts/custom-xml";
import type {
  BookmarkStartOptions,
  MarkupRangeOptions,
  MoveRangeStartOptions,
} from "@parts/paragraph/links/bookmark";
import type { ShadingProperties } from "@shared/shading";

import type { TableGridChangeOptions } from "./grid";
import type { TableCellSpacingProperties } from "./table-cell-spacing";
import type {
  TableBordersOptions,
  TableFloatOptions,
  TablePropertiesChangeOptions,
} from "./table-properties";
import type { TableCellMarginOptions } from "./table-properties/table-cell-margin";
import type { TableLayoutType } from "./table-properties/table-layout";
import type { TableLookOptions } from "./table-properties/table-look";
import type { TableJustification, TablePropertyOrder } from "./table-properties/table-properties";
import type { SdtRowOptions, TableRowOptions } from "./table-row";
import type { TableWidthProperties } from "./table-width";

/**
 * Options for creating a Table element.
 *
 * Avoid 0-width columns — they render incorrectly; the layout algorithm
 * expands columns to fit content even in 'auto' layout.
 */
export interface TableOptions extends BaseTableOptions<
  TableRowOptions | { sdt: SdtRowOptions } | { customXml: CustomXmlRowOptions } | TableMarkerChild
> {
  width?: TableWidthProperties;
  /** Count of legacy empty w:tblCol children when a producer omitted w:gridCol. */
  legacyColumnCount?: number;
  /** Source w:gridCol widths, including null for an omitted optional w:w; round-trip only. */
  columnWidthsRaw?: (number | PositiveUniversalMeasure | null)[];
  columnWidthsRevision?: TableGridChangeOptions;
  margins?: TableCellMarginOptions;
  indent?: TableWidthProperties;
  float?: TableFloatOptions;
  /** Column sizing: "autofit" let content resize columns, "fixed" honor column widths. */
  layout?: (typeof TableLayoutType)[keyof typeof TableLayoutType];
  /** Source child order for w:tblPr; round-trip only for non-XSD sequences. */
  propertyOrder?: TablePropertyOrder[];
  style?: string;
  borders?: TableBordersOptions;
  /** Justification (ST_JcTable): where the table sits — "start"/"center"/"end" or the legacy "left"/"right". */
  alignment?: TableJustification;
  visuallyRightToLeft?: boolean;
  cellSpacing?: TableCellSpacingProperties;
  styleRowBandSize?: number;
  styleColBandSize?: number;
  caption?: string;
  description?: string;
  revision?: TablePropertiesChangeOptions;
  /**
   * Explicit w:tblLook content (CT_TblLook). The base 6-flags above are the
   * authoring shorthand for the boolean attributes; this carries the full
   * element including the legacy w:val hex bitmask only old Word files write.
   */
  tableLook?: TableLookOptions;
  /** Table-level shading (w:shd). */
  shading?: ShadingProperties;
}

/**
 * Run-level markers anchored directly under `w:tbl`, between rows — bookmark
 * and revision ranges can span row boundaries and cannot live inside `w:tr`.
 */
export type TableMarkerChild =
  | { commentRangeStart: MarkupRangeOptions }
  | { commentRangeEnd: MarkupRangeOptions }
  | { bookmarkStart: BookmarkStartOptions }
  | { bookmarkEnd: MarkupRangeOptions }
  | { moveFromRangeStart: MoveRangeStartOptions }
  | { moveFromRangeEnd: MarkupRangeOptions }
  | { moveToRangeStart: MoveRangeStartOptions }
  | { moveToRangeEnd: MarkupRangeOptions };
