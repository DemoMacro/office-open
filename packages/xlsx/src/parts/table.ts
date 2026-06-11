/**
 * Table types and descriptor for SpreadsheetML documents.
 *
 * Implements CT_Table from sml.xsd (transitional schema).
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { findChild, attr, attrNum, textOf, escapeXml } from "@office-open/xml";

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

// ── Helper ──

function buildAttrs(attrsMap: Record<string, string | number | boolean | undefined>): string {
  const parts: string[] = [];
  for (const [k, v] of Object.entries(attrsMap)) {
    if (v === undefined) continue;
    parts.push(` ${k}="${typeof v === "string" ? escapeXml(v) : String(v)}"`);
  }
  return parts.join("");
}

// ── Descriptor ──

export const tableDesc: CustomDescriptor<TableOptions> = {
  kind: "custom",

  stringify(o, _ctx) {
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
        ` xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"${buildAttrs(rootAttrs)}>`,
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
        p.push(`<tableColumn${buildAttrs(colAttrs)}>${inner.join("")}</tableColumn>`);
      } else {
        p.push(`<tableColumn${buildAttrs(colAttrs)}/>`);
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
      p.push(`<tableStyleInfo${buildAttrs(styleAttrs)}/>`);
    } else {
      // Default style
      p.push(
        '<tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>',
      );
    }

    p.push("</table>");
    return p.join("");
  },

  parse(el, _ctx) {
    const result: Record<string, unknown> = {};

    // Root attributes
    const id = attrNum(el, "id");
    if (id !== undefined) result.id = id;
    if (attr(el, "name")) result.name = attr(el, "name");
    if (attr(el, "displayName")) result.displayName = attr(el, "displayName");
    if (attr(el, "ref")) result.ref = attr(el, "ref");
    const headerRowCount = attrNum(el, "headerRowCount");
    if (headerRowCount !== undefined) result.headerRowCount = headerRowCount;
    const totalsRowCount = attrNum(el, "totalsRowCount");
    if (totalsRowCount !== undefined) result.totalsRowCount = totalsRowCount;
    if (attr(el, "totalsRowShown") === "0") result.totalsRowShown = false;
    if (attr(el, "tableType")) result.tableType = attr(el, "tableType");
    if (attr(el, "insertRowShift") === "1") result.insertRowShift = true;
    if (attr(el, "published") === "1") result.published = true;

    // Auto filter
    const afEl = findChild(el, "autoFilter");
    if (afEl) result.autoFilter = attr(afEl, "ref") ?? "";

    // Table columns
    const colsEl = findChild(el, "tableColumns");
    if (colsEl) {
      const columns: Record<string, unknown>[] = [];
      for (const colEl of colsEl.elements ?? []) {
        if (colEl.name !== "tableColumn") continue;
        const col: Record<string, unknown> = {};
        const colId = attrNum(colEl, "id");
        if (colId !== undefined) col.id = colId;
        col.name = attr(colEl, "name") ?? "";
        if (attr(colEl, "totalsRowFunction"))
          col.totalsRowFunction = attr(colEl, "totalsRowFunction");
        if (attr(colEl, "totalsRowLabel")) col.totalsRowLabel = attr(colEl, "totalsRowLabel");
        const trfEl = findChild(colEl, "calculatedColumnFormula");
        if (trfEl) col.calculatedColumnFormula = textOf(trfEl);
        if (attr(colEl, "uniqueName")) col.uniqueName = attr(colEl, "uniqueName");
        const qtfId = attrNum(colEl, "queryTableFieldId");
        if (qtfId !== undefined) col.queryTableFieldId = qtfId;
        const hrDxfId = attrNum(colEl, "headerRowDxfId");
        if (hrDxfId !== undefined) col.headerRowDxfId = hrDxfId;
        const dDxfId = attrNum(colEl, "dataDxfId");
        if (dDxfId !== undefined) col.dataDxfId = dDxfId;
        const trDxfId = attrNum(colEl, "totalsRowDxfId");
        if (trDxfId !== undefined) col.totalsRowDxfId = trDxfId;
        if (attr(colEl, "headerRowCellStyle"))
          col.headerRowCellStyle = attr(colEl, "headerRowCellStyle");
        if (attr(colEl, "dataCellStyle")) col.dataCellStyle = attr(colEl, "dataCellStyle");
        if (attr(colEl, "totalsRowCellStyle"))
          col.totalsRowCellStyle = attr(colEl, "totalsRowCellStyle");
        columns.push(col);
      }
      result.columns = columns;
    }

    // Table style info
    const siEl = findChild(el, "tableStyleInfo");
    if (siEl) {
      const style: Record<string, unknown> = {};
      if (attr(siEl, "name")) style.name = attr(siEl, "name");
      if (attr(siEl, "showFirstColumn") === "1") style.showFirstColumn = true;
      if (attr(siEl, "showLastColumn") === "1") style.showLastColumn = true;
      if (attr(siEl, "showRowStripes") === "1") style.showRowStripes = true;
      if (attr(siEl, "showColumnStripes") === "1") style.showColumnStripes = true;
      result.style = style;
    }

    // Differential format IDs
    const hrDxfId = attrNum(el, "headerRowDxfId");
    if (hrDxfId !== undefined) result.headerRowDxfId = hrDxfId;
    const dDxfId = attrNum(el, "dataDxfId");
    if (dDxfId !== undefined) result.dataDxfId = dDxfId;
    const trDxfId = attrNum(el, "totalsRowDxfId");
    if (trDxfId !== undefined) result.totalsRowDxfId = trDxfId;
    const hrbDxfId = attrNum(el, "headerRowBorderDxfId");
    if (hrbDxfId !== undefined) result.headerRowBorderDxfId = hrbDxfId;
    const tbDxfId = attrNum(el, "tableBorderDxfId");
    if (tbDxfId !== undefined) result.tableBorderDxfId = tbDxfId;
    const trbDxfId = attrNum(el, "totalsRowBorderDxfId");
    if (trbDxfId !== undefined) result.totalsRowBorderDxfId = trbDxfId;
    if (attr(el, "headerRowCellStyle")) result.headerRowCellStyle = attr(el, "headerRowCellStyle");
    if (attr(el, "dataCellStyle")) result.dataCellStyle = attr(el, "dataCellStyle");
    if (attr(el, "totalsRowCellStyle")) result.totalsRowCellStyle = attr(el, "totalsRowCellStyle");

    return result as Record<string, unknown>;
  },
};
