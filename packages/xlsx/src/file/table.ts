/**
 * Table types for SpreadsheetML documents.
 *
 * Implements CT_Table from sml.xsd (transitional schema).
 *
 * @module
 */

// ── Totals row function (ST_TotalsRowFunction) ──

export const TotalsRowFunction = {
  NONE: "none",
  SUM: "sum",
  MIN: "min",
  MAX: "max",
  AVERAGE: "average",
  COUNT: "count",
  COUNT_NUMS: "countNums",
  STD_DEV: "stdDev",
  VAR: "var",
  CUSTOM: "custom",
} as const;

export type TotalsRowFunction = (typeof TotalsRowFunction)[keyof typeof TotalsRowFunction];

// ── Table type (ST_TableType) ──

export const TableType = {
  WORKSHEET: "worksheet",
  XML: "xml",
  QUERY_TABLE: "queryTable",
} as const;

export type TableType = (typeof TableType)[keyof typeof TableType];

// ── Options interfaces ──

export interface TableStyleInfoOptions {
  /** Table style name, e.g. "TableStyleMedium9" */
  name?: string;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
  showRowStripes?: boolean;
  showColumnStripes?: boolean;
}

export interface TableColumnOptions {
  /** Column name (used in header row) */
  name: string;
  /** Totals row function */
  totalsRowFunction?: TotalsRowFunction;
  /** Totals row label (used when totalsRowFunction is "none" or "custom") */
  totalsRowLabel?: string;
  /** Calculated column formula */
  calculatedColumnFormula?: string;
  /** Totals row formula (CT_TableColumn/totalsRowFormula, used when totalsRowFunction is "custom") */
  totalsRowFormula?: string;
  /** Whether totals row formula is array (CT_TableFormula @array) */
  totalsRowFormulaArray?: boolean;
  /** Whether calculated column formula is array (CT_TableFormula @array) */
  calculatedColumnFormulaArray?: boolean;
  /** Unique column name for structured references (CT_TableColumn @uniqueName) */
  uniqueName?: string;
  /** Query table field ID (CT_TableColumn @queryTableFieldId) */
  queryTableFieldId?: number;
  /** Header row differential format index */
  headerRowDxfId?: number;
  /** Data differential format index */
  dataDxfId?: number;
  /** Totals row differential format index */
  totalsRowDxfId?: number;
  /** Header row cell style name */
  headerRowCellStyle?: string;
  /** Data cell style name */
  dataCellStyle?: string;
  /** Totals row cell style name */
  totalsRowCellStyle?: string;
}

export interface TableOptions {
  /** Unique table id (1-based, must be unique across the workbook) */
  id: number;
  /** Table name (used in structured references) */
  name?: string;
  /** Display name (required by XSD, defaults to name if not set) */
  displayName: string;
  /** Data range, e.g. "A1:D10" */
  ref: string;
  /** Column definitions */
  columns: TableColumnOptions[];
  /** Number of header rows (default: 1) */
  headerRowCount?: number;
  /** Number of totals rows (default: 0) */
  totalsRowCount?: number;
  /** Whether to show totals row (default: true when totalsRowCount > 0) */
  totalsRowShown?: boolean;
  /** Table type (default: "worksheet") */
  tableType?: TableType;
  /** Table style */
  style?: TableStyleInfoOptions;
  /** Auto-filter reference (defaults to ref) */
  autoFilter?: string;
  /** Insert row shifts existing rows (CT_Table @insertRowShift) */
  insertRowShift?: boolean;
  /** Published to server (CT_Table @published) */
  published?: boolean;
  /** Header row differential format index */
  headerRowDxfId?: number;
  /** Data differential format index */
  dataDxfId?: number;
  /** Totals row differential format index */
  totalsRowDxfId?: number;
  /** Header row border differential format index */
  headerRowBorderDxfId?: number;
  /** Table border differential format index */
  tableBorderDxfId?: number;
  /** Totals row border differential format index */
  totalsRowBorderDxfId?: number;
  /** Header row cell style name */
  headerRowCellStyle?: string;
  /** Data cell style name */
  dataCellStyle?: string;
  /** Totals row cell style name */
  totalsRowCellStyle?: string;
}
