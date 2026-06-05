/**
 * Table XML generator — generates xl/tables/table{n}.xml.
 *
 * Implements CT_Table from sml.xsd (transitional schema).
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { escapeXml } from "@office-open/xml";

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
  readonly name?: string;
  readonly showFirstColumn?: boolean;
  readonly showLastColumn?: boolean;
  readonly showRowStripes?: boolean;
  readonly showColumnStripes?: boolean;
}

export interface TableColumnOptions {
  /** Column name (used in header row) */
  readonly name: string;
  /** Totals row function */
  readonly totalsRowFunction?: TotalsRowFunction;
  /** Totals row label (used when totalsRowFunction is "none" or "custom") */
  readonly totalsRowLabel?: string;
  /** Calculated column formula */
  readonly calculatedColumnFormula?: string;
  /** Totals row formula (CT_TableColumn/totalsRowFormula, used when totalsRowFunction is "custom") */
  readonly totalsRowFormula?: string;
  /** Whether totals row formula is array (CT_TableFormula @array) */
  readonly totalsRowFormulaArray?: boolean;
  /** Whether calculated column formula is array (CT_TableFormula @array) */
  readonly calculatedColumnFormulaArray?: boolean;
  /** Unique column name for structured references (CT_TableColumn @uniqueName) */
  readonly uniqueName?: string;
  /** Query table field ID (CT_TableColumn @queryTableFieldId) */
  readonly queryTableFieldId?: number;
  /** Header row differential format index */
  readonly headerRowDxfId?: number;
  /** Data differential format index */
  readonly dataDxfId?: number;
  /** Totals row differential format index */
  readonly totalsRowDxfId?: number;
  /** Header row cell style name */
  readonly headerRowCellStyle?: string;
  /** Data cell style name */
  readonly dataCellStyle?: string;
  /** Totals row cell style name */
  readonly totalsRowCellStyle?: string;
}

export interface TableOptions {
  /** Unique table id (1-based, must be unique across the workbook) */
  readonly id: number;
  /** Table name (used in structured references) */
  readonly name?: string;
  /** Display name (required by XSD, defaults to name if not set) */
  readonly displayName: string;
  /** Data range, e.g. "A1:D10" */
  readonly ref: string;
  /** Column definitions */
  readonly columns: readonly TableColumnOptions[];
  /** Number of header rows (default: 1) */
  readonly headerRowCount?: number;
  /** Number of totals rows (default: 0) */
  readonly totalsRowCount?: number;
  /** Whether to show totals row (default: true when totalsRowCount > 0) */
  readonly totalsRowShown?: boolean;
  /** Table type (default: "worksheet") */
  readonly tableType?: TableType;
  /** Table style */
  readonly style?: TableStyleInfoOptions;
  /** Auto-filter reference (defaults to ref) */
  readonly autoFilter?: string;
  /** Insert row shifts existing rows (CT_Table @insertRowShift) */
  readonly insertRowShift?: boolean;
  /** Published to server (CT_Table @published) */
  readonly published?: boolean;
  /** Header row differential format index */
  readonly headerRowDxfId?: number;
  /** Data differential format index */
  readonly dataDxfId?: number;
  /** Totals row differential format index */
  readonly totalsRowDxfId?: number;
  /** Header row border differential format index */
  readonly headerRowBorderDxfId?: number;
  /** Table border differential format index */
  readonly tableBorderDxfId?: number;
  /** Totals row border differential format index */
  readonly totalsRowBorderDxfId?: number;
  /** Header row cell style name */
  readonly headerRowCellStyle?: string;
  /** Data cell style name */
  readonly dataCellStyle?: string;
  /** Totals row cell style name */
  readonly totalsRowCellStyle?: string;
}

/**
 * Table XML component — generates xl/tables/table{n}.xml.
 *
 * Follows the zero-allocation string concatenation pattern used by other
 * XLSX components.
 */
export class TableXml extends BaseXmlComponent {
  private readonly opts: TableOptions;

  public constructor(options: TableOptions) {
    super("table");
    this.opts = options;
  }

  public override toXml(_context: Context): string {
    const o = this.opts;
    const p: string[] = [];

    // Root element with attributes
    const rootAttrs: Record<string, string | number | boolean | undefined> = {
      id: o.id,
      name: o.name ?? o.displayName,
      displayName: o.displayName,
      ref: o.ref,
    };
    if (o.tableType && o.tableType !== "worksheet") {
      rootAttrs.tableType = o.tableType;
    }
    if (o.headerRowCount !== undefined && o.headerRowCount !== 1) {
      rootAttrs.headerRowCount = o.headerRowCount;
    }
    if (o.totalsRowCount !== undefined && o.totalsRowCount > 0) {
      rootAttrs.totalsRowCount = o.totalsRowCount;
    }
    if (o.totalsRowShown === false) {
      rootAttrs.totalsRowShown = 0;
    }
    if (o.insertRowShift) rootAttrs.insertRowShift = 1;
    if (o.published) rootAttrs.published = 1;
    if (o.headerRowDxfId !== undefined) rootAttrs.headerRowDxfId = o.headerRowDxfId;
    if (o.dataDxfId !== undefined) rootAttrs.dataDxfId = o.dataDxfId;
    if (o.totalsRowDxfId !== undefined) rootAttrs.totalsRowDxfId = o.totalsRowDxfId;
    if (o.headerRowBorderDxfId !== undefined)
      rootAttrs.headerRowBorderDxfId = o.headerRowBorderDxfId;
    if (o.tableBorderDxfId !== undefined) rootAttrs.tableBorderDxfId = o.tableBorderDxfId;
    if (o.totalsRowBorderDxfId !== undefined)
      rootAttrs.totalsRowBorderDxfId = o.totalsRowBorderDxfId;
    if (o.headerRowCellStyle) rootAttrs.headerRowCellStyle = o.headerRowCellStyle;
    if (o.dataCellStyle) rootAttrs.dataCellStyle = o.dataCellStyle;
    if (o.totalsRowCellStyle) rootAttrs.totalsRowCellStyle = o.totalsRowCellStyle;

    p.push(
      `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` +
        ` xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"` +
        ` mc:Ignorable="xr xr2"` +
        ` xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"` +
        ` xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"${this.buildAttrs(rootAttrs)}>`,
    );

    // autoFilter (optional, before tableColumns per XSD sequence)
    if (o.autoFilter !== undefined) {
      p.push(`<autoFilter ref="${escapeXml(o.autoFilter)}"/>`);
    }

    // tableColumns (required)
    p.push(`<tableColumns count="${o.columns.length}">`);
    for (let i = 0; i < o.columns.length; i++) {
      const col = o.columns[i];
      const colAttrs: Record<string, string | number | boolean | undefined> = {
        id: i + 1,
        name: col.name,
      };

      const inner: string[] = [];

      // calculatedColumnFormula
      if (col.calculatedColumnFormula !== undefined) {
        const fAttrs = col.calculatedColumnFormulaArray ? ' array="1"' : "";
        inner.push(
          `<calculatedColumnFormula${fAttrs}>${escapeXml(col.calculatedColumnFormula)}</calculatedColumnFormula>`,
        );
      }

      // totalsRowFormula (when totalsRowFunction is "custom")
      if (col.totalsRowFormula !== undefined) {
        const fAttrs = col.totalsRowFormulaArray ? ' array="1"' : "";
        inner.push(
          `<totalsRowFormula${fAttrs}>${escapeXml(col.totalsRowFormula)}</totalsRowFormula>`,
        );
      }

      // totalsRowFormula (only when totalsRowFunction is "custom")
      if (col.totalsRowFunction === TotalsRowFunction.CUSTOM && col.totalsRowLabel) {
        // Use totalsRowLabel for custom display
      }

      if (col.totalsRowFunction !== undefined && col.totalsRowFunction !== TotalsRowFunction.NONE) {
        colAttrs.totalsRowFunction = col.totalsRowFunction;
      }
      if (col.totalsRowLabel !== undefined) {
        colAttrs.totalsRowLabel = col.totalsRowLabel;
      }
      if (col.uniqueName) colAttrs.uniqueName = col.uniqueName;
      if (col.queryTableFieldId !== undefined) colAttrs.queryTableFieldId = col.queryTableFieldId;
      if (col.headerRowDxfId !== undefined) colAttrs.headerRowDxfId = col.headerRowDxfId;
      if (col.dataDxfId !== undefined) colAttrs.dataDxfId = col.dataDxfId;
      if (col.totalsRowDxfId !== undefined) colAttrs.totalsRowDxfId = col.totalsRowDxfId;
      if (col.headerRowCellStyle) colAttrs.headerRowCellStyle = col.headerRowCellStyle;
      if (col.dataCellStyle) colAttrs.dataCellStyle = col.dataCellStyle;
      if (col.totalsRowCellStyle) colAttrs.totalsRowCellStyle = col.totalsRowCellStyle;

      if (inner.length > 0) {
        p.push(`<tableColumn${this.buildAttrs(colAttrs)}>${inner.join("")}</tableColumn>`);
      } else {
        p.push(`<tableColumn${this.buildAttrs(colAttrs)}/>`);
      }
    }
    p.push("</tableColumns>");

    // tableStyleInfo (optional)
    if (o.style) {
      const s = o.style;
      const styleAttrs: Record<string, string | number | boolean | undefined> = {};
      if (s.name !== undefined) styleAttrs.name = s.name;
      if (s.showFirstColumn) styleAttrs.showFirstColumn = 1;
      if (s.showLastColumn) styleAttrs.showLastColumn = 1;
      if (s.showRowStripes !== false) styleAttrs.showRowStripes = 1;
      if (s.showColumnStripes) styleAttrs.showColumnStripes = 1;
      p.push(`<tableStyleInfo${this.buildAttrs(styleAttrs)}/>`);
    } else {
      // Default style
      p.push(
        '<tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>',
      );
    }

    p.push("</table>");
    return p.join("");
  }

  private buildAttrs(attrs: Record<string, string | number | boolean | undefined>): string {
    const parts: string[] = [];
    for (const [k, v] of Object.entries(attrs)) {
      if (v === undefined) continue;
      parts.push(` ${k}="${typeof v === "string" ? escapeXml(v) : String(v)}"`);
    }
    return parts.join("");
  }
}
