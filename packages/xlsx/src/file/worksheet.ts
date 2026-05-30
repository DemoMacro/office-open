import type { SharedStrings } from "@file/shared-strings";
import type { Styles, StyleOptions } from "@file/styles";
/**
 * Worksheet component — generates xl/worksheets/sheet{n}.xml.
 *
 * @module
 */
import { IgnoreIfEmptyXmlComponent } from "@file/xml-components";
import type { Context, IXmlableObject } from "@file/xml-components";
import type { ChartSpaceOptions } from "@office-open/core";
import { attrs, escapeXml, selfCloseElement } from "@office-open/xml";

export interface ColumnOptions {
  readonly min: number;
  readonly max: number;
  readonly width?: number;
  readonly hidden?: boolean;
  readonly customWidth?: boolean;
}

export interface RowOptions {
  readonly cells?: readonly CellOptions[];
  readonly height?: number;
  readonly hidden?: boolean;
  readonly rowNumber?: number;
}

export interface CellOptions {
  readonly value?: string | number | boolean | Date | null;
  readonly reference?: string;
  /** Direct style index (for pre-resolved styles) */
  readonly styleIndex?: number;
  /** Style options (resolved to index at compile time) */
  readonly style?: StyleOptions;
  /** Formula options. When set, value becomes the cached result. */
  readonly formula?: FormulaOptions;
}

/** Cell formula type (maps to ST_CellFormulaType). */
export const FormulaType = {
  NORMAL: "normal",
  ARRAY: "array",
  SHARED: "shared",
} as const;

export type FormulaType = (typeof FormulaType)[keyof typeof FormulaType];

/** Options for a cell formula (maps to CT_CellFormula). */
export interface FormulaOptions {
  /** Formula expression, e.g. "SUM(A1:B1)" */
  readonly formula: string;
  /** Formula type (default: "normal") */
  readonly type?: FormulaType;
  /** Reference range for array/shared formulas, e.g. "C1:C10" */
  readonly reference?: string;
  /** Shared formula group index (required for shared formulas) */
  readonly sharedIndex?: number;
}

export interface MergeCellOptions {
  readonly from: { readonly row: number; readonly col: number };
  readonly to: { readonly row: number; readonly col: number };
}

export interface FreezePaneOptions {
  /** Row split position (1-based, freezes rows above) */
  readonly row?: number;
  /** Column split position (1-based, freezes columns to the left) */
  readonly col?: number;
}

export interface WorksheetImageOptions {
  readonly data: Uint8Array;
  readonly type: "png" | "jpeg" | "jpg";
  readonly col: number;
  readonly row: number;
}

export interface WorksheetChartOptions extends ChartSpaceOptions {
  /** 1-based column position for the chart */
  readonly col: number;
  /** 1-based row position for the chart */
  readonly row: number;
}

export type DataValidationType =
  | "none"
  | "whole"
  | "decimal"
  | "list"
  | "date"
  | "time"
  | "textLength"
  | "custom";
export type DataValidationOperator =
  | "between"
  | "notBetween"
  | "equal"
  | "notEqual"
  | "greaterThan"
  | "lessThan"
  | "greaterThanOrEqual"
  | "lessThanOrEqual";

export interface DataValidationOptions {
  /** Cell range, e.g. "A1:A10" */
  readonly sqref: string;
  readonly type?: DataValidationType;
  readonly operator?: DataValidationOperator;
  readonly formula1?: string;
  readonly formula2?: string;
  readonly allowBlank?: boolean;
  readonly showErrorMessage?: boolean;
  readonly errorTitle?: string;
  readonly error?: string;
  readonly showInputMessage?: boolean;
  readonly promptTitle?: string;
  readonly prompt?: string;
}

export type ConditionalFormatType =
  | "cellIs"
  | "containsText"
  | "expression"
  | "top10"
  | "aboveAverage";
export type ConditionalFormatOperator =
  | "lessThan"
  | "lessThanOrEqual"
  | "equal"
  | "notEqual"
  | "greaterThanOrEqual"
  | "greaterThan"
  | "between"
  | "notBetween"
  | "containsText"
  | "notContains"
  | "beginsWith"
  | "endsWith";

export interface ConditionalFormatRule {
  readonly type: ConditionalFormatType;
  readonly operator?: ConditionalFormatOperator;
  /** Formula(s) — up to 3 */
  readonly formulas?: readonly string[];
  readonly priority?: number;
  /** Reference to a dxf (differential format) in the styles table */
  readonly dxfId?: number;
}

export interface ConditionalFormatOptions {
  /** Cell range, e.g. "A1:A10" */
  readonly sqref: string;
  readonly rules: readonly ConditionalFormatRule[];
}

export interface WorksheetOptions {
  readonly name?: string;
  readonly children?: readonly RowOptions[];
  readonly columns?: readonly ColumnOptions[];
  readonly mergeCells?: readonly MergeCellOptions[];
  readonly freezePanes?: FreezePaneOptions;
  /** Auto-filter range, e.g. "A1:D10" */
  readonly autoFilter?: string;
  readonly images?: readonly WorksheetImageOptions[];
  readonly charts?: readonly WorksheetChartOptions[];
  readonly dataValidations?: readonly DataValidationOptions[];
  readonly conditionalFormats?: readonly ConditionalFormatOptions[];
}

export class Worksheet extends IgnoreIfEmptyXmlComponent {
  private readonly rows: readonly RowOptions[];
  private readonly columns: readonly ColumnOptions[];
  private readonly mergeCells: readonly MergeCellOptions[];
  private readonly freezePanes?: FreezePaneOptions;
  private readonly autoFilter?: string;
  private readonly images: readonly WorksheetImageOptions[];
  private readonly chartOptions: readonly WorksheetChartOptions[];
  private readonly dataValidations: readonly DataValidationOptions[];
  private readonly conditionalFormats: readonly ConditionalFormatOptions[];

  public constructor(options: WorksheetOptions) {
    super("worksheet");
    this.rows = options.children ?? [];
    this.columns = options.columns ?? [];
    this.mergeCells = options.mergeCells ?? [];
    this.freezePanes = options.freezePanes;
    this.autoFilter = options.autoFilter;
    this.images = options.images ?? [];
    this.chartOptions = options.charts ?? [];
    this.dataValidations = options.dataValidations ?? [];
    this.conditionalFormats = options.conditionalFormats ?? [];
  }

  public get imageOptions(): readonly WorksheetImageOptions[] {
    return this.images;
  }

  public get charts(): readonly WorksheetChartOptions[] {
    return this.chartOptions;
  }

  public override prepForXml(context: Context): IXmlableObject | undefined {
    const fileData = context.fileData as
      | { sharedStrings?: SharedStrings; styles?: Styles }
      | undefined;
    const sharedStrings = fileData?.sharedStrings;
    const styles = fileData?.styles;

    const children: IXmlableObject[] = [
      {
        _attr: {
          xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
      },
    ];

    // Dimension — defines the used range of the sheet
    const maxRow = this.rows.length;
    let maxCol = 0;
    for (const row of this.rows) {
      if (row.cells && row.cells.length > maxCol) maxCol = row.cells.length;
    }
    if (maxRow > 0 && maxCol > 0) {
      const dimRef = `A1:${this.defaultCellRef(maxRow, maxCol)}`;
      children.push({ dimension: { _attr: { ref: dimRef } } });
    }

    // Sheet views (freeze panes) — must come before cols per XSD sequence
    if (this.freezePanes) {
      const fp = this.freezePanes;
      const ySplit = fp.row ? fp.row : 0;
      const xSplit = fp.col ? fp.col : 0;
      const topRow = fp.row ? fp.row + 1 : 1;
      const leftCol = fp.col ? fp.col + 1 : 1;
      const topLeftCell = this.defaultCellRef(topRow, leftCol);
      const attr: Record<string, string | number> = {
        ySplit,
        xSplit,
        topLeftCell,
        activePane:
          ySplit > 0 && xSplit > 0 ? "bottomRight" : ySplit > 0 ? "bottomLeft" : "topRight",
        state: "frozen",
      };
      children.push({
        sheetViews: [
          {
            sheetView: [
              { _attr: { tabSelected: 1, workbookViewId: 0 } },
              { pane: { _attr: attr } },
            ],
          },
        ],
      });
    } else {
      children.push({
        sheetViews: [{ sheetView: [{ _attr: { tabSelected: 1, workbookViewId: 0 } }] }],
      });
    }

    children.push({ sheetFormatPr: { _attr: { defaultRowHeight: 15 } } });

    // Column definitions (after sheetViews, before sheetData)
    if (this.columns.length > 0) {
      const colElements: IXmlableObject[] = [];
      for (const col of this.columns) {
        const colAttrs: Record<string, string | number> = {
          min: col.min,
          max: col.max,
        };
        if (col.width !== undefined) {
          colAttrs.width = col.width;
          colAttrs.customWidth = 1;
        }
        if (col.hidden) {
          colAttrs.hidden = 1;
        }
        colElements.push({ col: { _attr: colAttrs } });
      }
      children.push({ cols: colElements });
    }

    // Sheet data (rows + cells)
    const sheetDataChildren: IXmlableObject[] = [];
    for (let i = 0; i < this.rows.length; i++) {
      const rowOpts = this.rows[i];
      const rowNumber = rowOpts.rowNumber ?? i + 1;
      const rowAttrs: Record<string, string | number> = { r: rowNumber };
      if (rowOpts.height !== undefined) {
        rowAttrs.ht = rowOpts.height;
        rowAttrs.customHeight = 1;
      }
      if (rowOpts.hidden) {
        rowAttrs.hidden = 1;
      }

      const cellElements: IXmlableObject[] = [];
      if (rowOpts.cells) {
        for (let j = 0; j < rowOpts.cells.length; j++) {
          const cell = rowOpts.cells[j];
          const ref = cell.reference ?? this.defaultCellRef(rowNumber, j + 1);
          const cellObj = this.buildCell(ref, cell, sharedStrings, styles);
          if (cellObj) cellElements.push(cellObj);
        }
      }

      sheetDataChildren.push({ row: [{ _attr: rowAttrs }, ...cellElements] });
    }

    children.push({ sheetData: sheetDataChildren });

    // Auto filter (after sheetData, before mergeCells — per XSD sequence)
    if (this.autoFilter) {
      children.push({ autoFilter: { _attr: { ref: this.autoFilter } } });
    }

    // Merge cells (after autoFilter)
    if (this.mergeCells.length > 0) {
      const mergeElements: IXmlableObject[] = [{ _attr: { count: this.mergeCells.length } }];
      for (const mc of this.mergeCells) {
        const fromRef = this.defaultCellRef(mc.from.row, mc.from.col);
        const toRef = this.defaultCellRef(mc.to.row, mc.to.col);
        mergeElements.push({
          mergeCell: { _attr: { ref: `${fromRef}:${toRef}` } },
        });
      }
      children.push({ mergeCells: mergeElements });
    }

    // Conditional formatting (after autoFilter, before dataValidations)
    if (this.conditionalFormats.length > 0) {
      for (const cf of this.conditionalFormats) {
        const rules: IXmlableObject[] = [];
        for (let ri = 0; ri < cf.rules.length; ri++) {
          const rule = cf.rules[ri];
          const ruleAttrs: Record<string, string | number> = {
            type: rule.type,
            priority: rule.priority ?? ri + 1,
          };
          if (rule.operator) ruleAttrs.operator = rule.operator;
          if (rule.dxfId !== undefined) ruleAttrs.dxfId = rule.dxfId;

          const ruleChildren: IXmlableObject[] = [{ _attr: ruleAttrs }];

          if (rule.formulas) {
            for (const f of rule.formulas) {
              ruleChildren.push({ formula: [f] });
            }
          }
          rules.push({ cfRule: ruleChildren });
        }
        children.push({ conditionalFormatting: [{ _attr: { sqref: cf.sqref } }, ...rules] });
      }
    }

    // Data validations (after conditionalFormatting)
    if (this.dataValidations.length > 0) {
      const dvElements: IXmlableObject[] = [{ _attr: { count: this.dataValidations.length } }];
      for (const dv of this.dataValidations) {
        const dvAttrs: Record<string, string | number> = { sqref: dv.sqref };
        if (dv.type && dv.type !== "none") dvAttrs.type = dv.type;
        if (dv.operator) dvAttrs.operator = dv.operator;
        if (dv.allowBlank) dvAttrs.allowBlank = 1;
        if (dv.showErrorMessage) dvAttrs.showErrorMessage = 1;
        if (dv.showInputMessage) dvAttrs.showInputMessage = 1;
        if (dv.errorTitle) dvAttrs.errorTitle = dv.errorTitle;
        if (dv.error) dvAttrs.error = dv.error;
        if (dv.promptTitle) dvAttrs.promptTitle = dv.promptTitle;
        if (dv.prompt) dvAttrs.prompt = dv.prompt;

        const dvChildren: IXmlableObject[] = [{ _attr: dvAttrs }];
        if (dv.formula1 !== undefined) dvChildren.push({ formula1: [dv.formula1] });
        if (dv.formula2 !== undefined) dvChildren.push({ formula2: [dv.formula2] });
        dvElements.push({ dataValidation: dvChildren });
      }
      children.push({ dataValidations: dvElements });
    }

    children.push({
      pageMargins: {
        _attr: { left: 0.75, right: 0.75, top: 1, bottom: 1, header: 0.5, footer: 0.5 },
      },
    });

    return { worksheet: children };
  }

  /**
   * Zero-allocation fast path: directly concatenate XML string.
   * Bypasses the IXmlableObject intermediate tree entirely.
   */
  public override toXml(context: Context): string {
    const fileData = context.fileData as
      | { sharedStrings?: SharedStrings; styles?: Styles }
      | undefined;
    const sharedStrings = fileData?.sharedStrings;
    const styles = fileData?.styles;

    const p: string[] = [
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    ];

    // Dimension — defines the used range of the sheet
    const maxRow = this.rows.length;
    let maxCol = 0;
    for (const row of this.rows) {
      if (row.cells && row.cells.length > maxCol) maxCol = row.cells.length;
    }
    if (maxRow > 0 && maxCol > 0) {
      const dimRef = `A1:${this.defaultCellRef(maxRow, maxCol)}`;
      p.push(`<dimension ref="${dimRef}"/>`);
    }

    // Sheet views
    if (this.freezePanes) {
      const fp = this.freezePanes;
      const ySplit = fp.row ? fp.row : 0;
      const xSplit = fp.col ? fp.col : 0;
      const topRow = fp.row ? fp.row + 1 : 1;
      const leftCol = fp.col ? fp.col + 1 : 1;
      const topLeftCell = this.defaultCellRef(topRow, leftCol);
      const activePane =
        ySplit > 0 && xSplit > 0 ? "bottomRight" : ySplit > 0 ? "bottomLeft" : "topRight";
      p.push(
        '<sheetViews><sheetView tabSelected="1" workbookViewId="0">',
        `<pane ySplit="${ySplit}" xSplit="${xSplit}" topLeftCell="${topLeftCell}" activePane="${activePane}" state="frozen"/>`,
        "</sheetView></sheetViews>",
      );
    } else {
      // Default sheetView (required by Excel)
      p.push('<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>');
    }

    // Sheet format — default row height
    p.push('<sheetFormatPr defaultRowHeight="15"/>');

    // Column definitions
    if (this.columns.length > 0) {
      p.push("<cols>");
      for (const col of this.columns) {
        const colAttrs: Record<string, string | number | boolean | undefined> = {
          min: col.min,
          max: col.max,
        };
        if (col.width !== undefined) {
          colAttrs.width = col.width;
          colAttrs.customWidth = 1;
        }
        if (col.hidden) {
          colAttrs.hidden = 1;
        }
        p.push(selfCloseElement("col", attrs(colAttrs)));
      }
      p.push("</cols>");
    }

    // Sheet data (rows + cells) — the hot path
    p.push("<sheetData>");
    for (let i = 0; i < this.rows.length; i++) {
      const rowOpts = this.rows[i];
      const rowNumber = rowOpts.rowNumber ?? i + 1;
      const rowAttrs: Record<string, string | number | boolean | undefined> = { r: rowNumber };
      if (rowOpts.height !== undefined) {
        rowAttrs.ht = rowOpts.height;
        rowAttrs.customHeight = 1;
      }
      if (rowOpts.hidden) {
        rowAttrs.hidden = 1;
      }

      if (rowOpts.cells) {
        const rowParts: string[] = [];
        for (let j = 0; j < rowOpts.cells.length; j++) {
          const cell = rowOpts.cells[j];
          const ref = cell.reference ?? this.defaultCellRef(rowNumber, j + 1);
          const cellStr = this.buildCellString(ref, cell, sharedStrings, styles);
          if (cellStr) rowParts.push(cellStr);
        }
        p.push(`<row${attrs(rowAttrs)}>`, ...rowParts, "</row>");
      } else {
        p.push(`<row${attrs(rowAttrs)}/>`);
      }
    }
    p.push("</sheetData>");

    // Auto filter
    if (this.autoFilter) {
      p.push(selfCloseElement("autoFilter", attrs({ ref: this.autoFilter })));
    }

    // Merge cells
    if (this.mergeCells.length > 0) {
      p.push(`<mergeCells count="${this.mergeCells.length}">`);
      for (const mc of this.mergeCells) {
        const fromRef = this.defaultCellRef(mc.from.row, mc.from.col);
        const toRef = this.defaultCellRef(mc.to.row, mc.to.col);
        p.push(selfCloseElement("mergeCell", attrs({ ref: `${fromRef}:${toRef}` })));
      }
      p.push("</mergeCells>");
    }

    // Conditional formatting
    if (this.conditionalFormats.length > 0) {
      for (const cf of this.conditionalFormats) {
        p.push(`<conditionalFormatting sqref="${cf.sqref}">`);
        for (let ri = 0; ri < cf.rules.length; ri++) {
          const rule = cf.rules[ri];
          const ruleAttrs: Record<string, string | number | boolean | undefined> = {
            type: rule.type,
            priority: rule.priority ?? ri + 1,
          };
          if (rule.operator) ruleAttrs.operator = rule.operator;
          if (rule.dxfId !== undefined) ruleAttrs.dxfId = rule.dxfId;
          if (rule.formulas && rule.formulas.length > 0) {
            const formulaParts = rule.formulas.map((f) => `<formula>${escapeXml(f)}</formula>`);
            p.push(`<cfRule${attrs(ruleAttrs)}>`, ...formulaParts, "</cfRule>");
          } else {
            p.push(selfCloseElement("cfRule", attrs(ruleAttrs)));
          }
        }
        p.push("</conditionalFormatting>");
      }
    }

    // Data validations
    if (this.dataValidations.length > 0) {
      p.push(`<dataValidations count="${this.dataValidations.length}">`);
      for (const dv of this.dataValidations) {
        const dvAttrs: Record<string, string | number | boolean | undefined> = { sqref: dv.sqref };
        if (dv.type && dv.type !== "none") dvAttrs.type = dv.type;
        if (dv.operator) dvAttrs.operator = dv.operator;
        if (dv.allowBlank) dvAttrs.allowBlank = 1;
        if (dv.showErrorMessage) dvAttrs.showErrorMessage = 1;
        if (dv.showInputMessage) dvAttrs.showInputMessage = 1;
        if (dv.errorTitle) dvAttrs.errorTitle = dv.errorTitle;
        if (dv.error) dvAttrs.error = dv.error;
        if (dv.promptTitle) dvAttrs.promptTitle = dv.promptTitle;
        if (dv.prompt) dvAttrs.prompt = dv.prompt;
        const inner: string[] = [];
        if (dv.formula1 !== undefined) inner.push(`<formula1>${escapeXml(dv.formula1)}</formula1>`);
        if (dv.formula2 !== undefined) inner.push(`<formula2>${escapeXml(dv.formula2)}</formula2>`);
        if (inner.length > 0) {
          p.push(`<dataValidation${attrs(dvAttrs)}>`, ...inner, "</dataValidation>");
        } else {
          p.push(selfCloseElement("dataValidation", attrs(dvAttrs)));
        }
      }
      p.push("</dataValidations>");
    }

    p.push('<pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>');
    p.push("</worksheet>");
    return p.join("");
  }

  /**
   * Build the <f> element string for a cell formula.
   */
  private buildFormulaString(fOpts: FormulaOptions): string {
    const fAttrs: Record<string, string | number | boolean | undefined> = {};
    if (fOpts.type && fOpts.type !== FormulaType.NORMAL) fAttrs.t = fOpts.type;
    if (fOpts.reference) fAttrs.ref = fOpts.reference;
    if (fOpts.sharedIndex !== undefined) fAttrs.si = fOpts.sharedIndex;

    const hasContent = fOpts.formula !== undefined && fOpts.formula !== "";

    if (hasContent) {
      return `<f${attrs(fAttrs)}>${escapeXml(fOpts.formula)}</f>`;
    }
    if (Object.keys(fAttrs).length > 0) {
      return selfCloseElement("f", attrs(fAttrs));
    }
    return "";
  }

  /**
   * Direct string serialization of a single cell — zero intermediate objects.
   */
  private buildCellString(
    ref: string,
    cell: CellOptions,
    sharedStrings?: SharedStrings,
    styles?: Styles,
  ): string {
    const cellAttrs: Record<string, string | number | boolean | undefined> = { r: ref };

    // Resolve style
    if (cell.style !== undefined && styles) {
      cellAttrs.s = styles.register(cell.style);
    } else if (cell.styleIndex !== undefined) {
      cellAttrs.s = cell.styleIndex;
    }

    const value = cell.value;

    // Formula path — formula takes precedence; value is the cached result.
    if (cell.formula) {
      const fStr = this.buildFormulaString(cell.formula);
      let vStr = "";
      if (value === null || value === undefined) {
        return `<c${attrs(cellAttrs)}>${fStr}</c>`;
      }
      if (typeof value === "number") {
        vStr = `<v>${value}</v>`;
      } else if (typeof value === "boolean") {
        cellAttrs.t = "b";
        vStr = `<v>${value ? 1 : 0}</v>`;
      } else if (typeof value === "string") {
        cellAttrs.t = "str";
        vStr = `<v>${escapeXml(value)}</v>`;
      } else if (value instanceof Date) {
        vStr = `<v>${this.dateToSerialNumber(value)}</v>`;
      }
      if (vStr) {
        return `<c${attrs(cellAttrs)}>${fStr}${vStr}</c>`;
      }
      return `<c${attrs(cellAttrs)}>${fStr}</c>`;
    }

    if (value === null || value === undefined) {
      if (cell.styleIndex !== undefined) {
        return selfCloseElement("c", attrs(cellAttrs));
      }
      return "";
    }

    if (typeof value === "string") {
      if (sharedStrings) {
        cellAttrs.t = "s";
        const idx = sharedStrings.register(value);
        return `<c${attrs(cellAttrs)}><v>${idx}</v></c>`;
      }
      cellAttrs.t = "inlineStr";
      return `<c${attrs(cellAttrs)}><is><t>${escapeXml(value)}</t></is></c>`;
    }

    if (typeof value === "number") {
      return `<c${attrs(cellAttrs)}><v>${value}</v></c>`;
    }

    if (typeof value === "boolean") {
      cellAttrs.t = "b";
      return `<c${attrs(cellAttrs)}><v>${value ? 1 : 0}</v></c>`;
    }

    if (value instanceof Date) {
      const serial = this.dateToSerialNumber(value);
      return `<c${attrs(cellAttrs)}><v>${serial}</v></c>`;
    }

    return "";
  }

  private defaultCellRef(row: number, col: number): string {
    return this.columnToLetter(col) + row;
  }

  private columnToLetter(col: number): string {
    let result = "";
    let n = col;
    while (n > 0) {
      const remainder = (n - 1) % 26;
      result = String.fromCharCode(65 + remainder) + result;
      n = Math.floor((n - 1) / 26);
    }
    return result;
  }

  private buildCell(
    ref: string,
    cell: CellOptions,
    sharedStrings?: SharedStrings,
    styles?: Styles,
  ): IXmlableObject | undefined {
    const attrs: Record<string, string | number> = { r: ref };

    // Resolve style: either direct index or register options
    if (cell.style !== undefined && styles) {
      attrs.s = styles.register(cell.style);
    } else if (cell.styleIndex !== undefined) {
      attrs.s = cell.styleIndex;
    }

    const value = cell.value;

    // Formula path — formula takes precedence; value is the cached result.
    if (cell.formula) {
      const children: IXmlableObject[] = [{ _attr: attrs }];

      const fOpts = cell.formula;
      const fAttrs: Record<string, string | number> = {};
      if (fOpts.type && fOpts.type !== FormulaType.NORMAL) fAttrs.t = fOpts.type;
      if (fOpts.reference) fAttrs.ref = fOpts.reference;
      if (fOpts.sharedIndex !== undefined) fAttrs.si = fOpts.sharedIndex;

      const hasContent = fOpts.formula !== undefined && fOpts.formula !== "";
      if (hasContent) {
        children.push({
          f: [Object.keys(fAttrs).length > 0 ? { _attr: fAttrs } : undefined, fOpts.formula].filter(
            Boolean,
          ) as IXmlableObject[],
        });
      } else if (Object.keys(fAttrs).length > 0) {
        children.push({ f: [{ _attr: fAttrs }] });
      }

      if (value !== null && value !== undefined) {
        if (typeof value === "number") {
          children.push({ v: [value] });
        } else if (typeof value === "boolean") {
          attrs.t = "b";
          children.push({ v: [value ? 1 : 0] });
        } else if (typeof value === "string") {
          attrs.t = "str";
          children.push({ v: [value] });
        } else if (value instanceof Date) {
          children.push({ v: [this.dateToSerialNumber(value)] });
        }
      }

      return { c: children };
    }

    if (value === null || value === undefined) {
      if (cell.styleIndex !== undefined) {
        return { c: [{ _attr: attrs }] };
      }
      return undefined;
    }

    if (typeof value === "string") {
      if (sharedStrings) {
        attrs.t = "s";
        const idx = sharedStrings.register(value);
        return { c: [{ _attr: attrs }, { v: [idx] }] };
      }
      // Fallback: inline string
      attrs.t = "inlineStr";
      return { c: [{ _attr: attrs }, { is: [{ t: [value] }] }] };
    }

    if (typeof value === "number") {
      return { c: [{ _attr: attrs }, { v: [value] }] };
    }

    if (typeof value === "boolean") {
      attrs.t = "b";
      return { c: [{ _attr: attrs }, { v: [value ? 1 : 0] }] };
    }

    if (value instanceof Date) {
      const serial = this.dateToSerialNumber(value);
      return { c: [{ _attr: attrs }, { v: [serial] }] };
    }

    return undefined;
  }

  private dateToSerialNumber(date: Date): number {
    const epoch = new Date(1899, 11, 30);
    const msPerDay = 86400000;
    return (date.getTime() - epoch.getTime()) / msPerDay;
  }
}
