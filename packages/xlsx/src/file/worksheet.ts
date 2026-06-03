import type { SharedStrings } from "@file/shared-strings";
import type { Styles, StyleOptions } from "@file/styles";
/**
 * Worksheet component — generates xl/worksheets/sheet{n}.xml.
 *
 * @module
 */
import { IgnoreIfEmptyXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import type { ChartSpaceOptions } from "@office-open/core";
import { attrs, escapeXml, selfCloseElement } from "@office-open/xml";

import type { PivotTableOptions } from "./pivot";
import type { TableOptions } from "./table";

export interface ColumnOptions {
  readonly min: number;
  readonly max: number;
  readonly width?: number;
  readonly hidden?: boolean;
  readonly customWidth?: boolean;
  readonly outlineLevel?: number;
  readonly collapsed?: boolean;
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

export interface SheetProtectionOptions {
  /** Plain-text password — legacy Excel hash is computed automatically */
  readonly password?: string;
  /** Modern encryption: algorithm name (e.g. "SHA-512") */
  readonly algorithmName?: string;
  /** Modern encryption: base64-encoded hash value */
  readonly hashValue?: string;
  /** Modern encryption: base64-encoded salt value */
  readonly saltValue?: string;
  /** Modern encryption: spin count for hash iteration */
  readonly spinCount?: number;
  /** Set true to enable sheet protection (required for protection flags to take effect) */
  readonly sheet?: boolean;
  readonly objects?: boolean;
  readonly scenarios?: boolean;
  readonly formatCells?: boolean;
  readonly formatColumns?: boolean;
  readonly formatRows?: boolean;
  readonly insertColumns?: boolean;
  readonly insertRows?: boolean;
  readonly insertHyperlinks?: boolean;
  readonly deleteColumns?: boolean;
  readonly deleteRows?: boolean;
  readonly selectLockedCells?: boolean;
  readonly sort?: boolean;
  readonly autoFilter?: boolean;
  readonly pivotTables?: boolean;
  readonly selectUnlockedCells?: boolean;
}

/** A named protected range within a sheet (CT_ProtectedRange) */
export interface ProtectedRangeOptions {
  /** Range reference (required), e.g. "A1:C10" */
  readonly sqref: string;
  /** Range name (required) */
  readonly name: string;
  /** Plain-text password — legacy hash computed automatically */
  readonly password?: string;
  /** Modern encryption: algorithm name */
  readonly algorithmName?: string;
  /** Modern encryption: base64-encoded hash value */
  readonly hashValue?: string;
  /** Modern encryption: base64-encoded salt value */
  readonly saltValue?: string;
  /** Modern encryption: spin count */
  readonly spinCount?: number;
  /** Security descriptor (SID string) */
  readonly securityDescriptor?: string;
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

export interface SheetViewOptions {
  readonly showGridLines?: boolean;
  readonly showRowColHeaders?: boolean;
  readonly showZeros?: boolean;
  readonly zoomScale?: number;
  readonly tabSelected?: boolean;
  readonly rightToLeft?: boolean;
}

export type HyperlinkTarget =
  | { readonly type: "external"; readonly url: string }
  | { readonly type: "internal"; readonly location: string };

export interface HyperlinkOptions {
  /** Cell reference, e.g. "A1" */
  readonly cell: string;
  /** Hyperlink target */
  readonly target: HyperlinkTarget;
  /** Tooltip text */
  readonly tooltip?: string;
  /** Display text */
  readonly display?: string;
}

export interface HeaderFooterOptions {
  readonly oddHeader?: string;
  readonly oddFooter?: string;
  readonly evenHeader?: string;
  readonly evenFooter?: string;
  readonly firstHeader?: string;
  readonly firstFooter?: string;
  readonly differentOddEven?: boolean;
  readonly differentFirst?: boolean;
}

export type PageOrientation = "default" | "portrait" | "landscape";

export interface PageSetupOptions {
  readonly paperSize?: number;
  readonly orientation?: PageOrientation;
  readonly scale?: number;
  readonly fitToWidth?: number;
  readonly fitToHeight?: number;
  readonly pageOrder?: "downThenOver" | "overThenDown";
  readonly useFirstPageNumber?: boolean;
  readonly firstPageNumber?: number;
}

export interface TabColorOptions {
  /** RGB color string, e.g. "FF0000" */
  readonly rgb?: string;
  /** Theme color index (0-based) */
  readonly theme?: number;
  /** Tint value (-1.0 to 1.0) */
  readonly tint?: number;
}

export interface CommentOptions {
  /** Cell reference, e.g. "A1" */
  readonly cell: string;
  /** Author name */
  readonly author: string;
  /** Comment text */
  readonly text: string;
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
  | "aboveAverage"
  | "colorScale"
  | "dataBar"
  | "iconSet";
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

/** Conditional format value object type (ST_CfvoType) */
export type CfvoType = "num" | "percent" | "max" | "min" | "formula" | "percentile";

/** Conditional format value object */
export interface CfvoOptions {
  readonly type: CfvoType;
  readonly val?: string | number;
  /** Greater than or equal (default: true) */
  readonly gte?: boolean;
}

/** Icon set type (ST_IconSetType) */
export type IconSetType =
  | "3Arrows"
  | "3ArrowsGray"
  | "3Flags"
  | "3TrafficLights1"
  | "3TrafficLights2"
  | "3Signs"
  | "3Symbols"
  | "3Symbols2"
  | "4Arrows"
  | "4ArrowsGray"
  | "4RedToBlack"
  | "4Rating"
  | "4TrafficLights"
  | "5Arrows"
  | "5ArrowsGray"
  | "5Rating"
  | "5Quarters";

/** Color scale rule configuration */
export interface ColorScaleOptions {
  /** Conditional format values (minimum 2, typically 2 or 3) */
  readonly cfvo: readonly CfvoOptions[];
  /** Colors for each value (same count as cfvo) — RGB hex without alpha, e.g. "FF0000" */
  readonly colors: readonly string[];
}

/** Data bar rule configuration */
export interface DataBarOptions {
  /** Minimum and maximum value objects (exactly 2) */
  readonly cfvo: readonly [CfvoOptions, CfvoOptions];
  /** Bar color — RGB hex without alpha, e.g. "638EC6" */
  readonly color: string;
  /** Minimum bar length as percentage (default: 10) */
  readonly minLength?: number;
  /** Maximum bar length as percentage (default: 90) */
  readonly maxLength?: number;
  /** Whether to show cell values (default: true) */
  readonly showValue?: boolean;
}

/** Icon set rule configuration */
export interface IconSetOptions {
  /** Conditional format values (minimum 2) */
  readonly cfvo: readonly CfvoOptions[];
  /** Icon set type (default: "3TrafficLights1") */
  readonly iconSet?: IconSetType;
  /** Whether to show cell values (default: true) */
  readonly showValue?: boolean;
  /** Whether values are percentages (default: true) */
  readonly percent?: boolean;
  /** Whether to reverse icon order (default: false) */
  readonly reverse?: boolean;
}

export interface ConditionalFormatRule {
  readonly type: ConditionalFormatType;
  readonly operator?: ConditionalFormatOperator;
  /** Formula(s) — up to 3 */
  readonly formulas?: readonly string[];
  readonly priority?: number;
  /** Reference to a dxf (differential format) in the styles table */
  readonly dxfId?: number;
  /** Color scale configuration (when type is "colorScale") */
  readonly colorScale?: ColorScaleOptions;
  /** Data bar configuration (when type is "dataBar") */
  readonly dataBar?: DataBarOptions;
  /** Icon set configuration (when type is "iconSet") */
  readonly iconSet?: IconSetOptions;
}

export interface ConditionalFormatOptions {
  /** Cell range, e.g. "A1:A10" */
  readonly sqref: string;
  readonly rules: readonly ConditionalFormatRule[];
}

export interface Top10FilterOptions {
  readonly colId: number;
  readonly top?: boolean;
  readonly percent?: boolean;
  readonly val: number;
}

export interface CustomFilterOptions {
  readonly colId: number;
  readonly operator?:
    | "equal"
    | "notEqual"
    | "greaterThan"
    | "greaterThanOrEqual"
    | "lessThan"
    | "lessThanOrEqual";
  readonly val?: string;
  readonly and?: boolean;
  readonly val2?: string;
}

export interface SortCondition {
  /** Cell reference for the sort column, e.g. "B1" */
  readonly ref: string;
  readonly descending?: boolean;
}

export interface AutoFilterOptions {
  /** Range, e.g. "A1:D10" */
  readonly ref: string;
  readonly top10?: readonly Top10FilterOptions[];
  readonly customFilters?: readonly CustomFilterOptions[];
  readonly sort?: readonly SortCondition[];
}

export interface WorksheetOptions {
  readonly name?: string;
  readonly rows?: readonly RowOptions[];
  readonly columns?: readonly ColumnOptions[];
  readonly mergeCells?: readonly MergeCellOptions[];
  readonly freezePanes?: FreezePaneOptions;
  readonly protection?: SheetProtectionOptions;
  /** Named protected ranges within this sheet */
  readonly protectedRanges?: readonly ProtectedRangeOptions[];
  /** Auto-filter configuration */
  readonly autoFilter?: string | AutoFilterOptions;
  readonly images?: readonly WorksheetImageOptions[];
  readonly charts?: readonly WorksheetChartOptions[];
  readonly dataValidations?: readonly DataValidationOptions[];
  readonly conditionalFormats?: readonly ConditionalFormatOptions[];
  readonly hyperlinks?: readonly HyperlinkOptions[];
  readonly comments?: readonly CommentOptions[];
  readonly headerFooter?: HeaderFooterOptions;
  readonly pageSetup?: PageSetupOptions;
  readonly tabColor?: TabColorOptions;
  readonly sheetView?: SheetViewOptions;
  readonly pivotTables?: readonly PivotTableOptions[];
  /** Tables (list objects) for this worksheet */
  readonly tables?: readonly TableOptions[];
}

export class Worksheet extends IgnoreIfEmptyXmlComponent {
  private readonly rows: readonly RowOptions[];
  private readonly columns: readonly ColumnOptions[];
  private readonly mergeCells: readonly MergeCellOptions[];
  private readonly freezePanes?: FreezePaneOptions;
  private readonly protection?: SheetProtectionOptions;
  private readonly protectedRanges: readonly ProtectedRangeOptions[];
  private readonly autoFilter?: string | AutoFilterOptions;
  private readonly images: readonly WorksheetImageOptions[];
  private readonly chartOptions: readonly WorksheetChartOptions[];
  private readonly dataValidations: readonly DataValidationOptions[];
  private readonly conditionalFormats: readonly ConditionalFormatOptions[];
  private readonly hyperlinks: readonly HyperlinkOptions[];
  private readonly comments: readonly CommentOptions[];
  private readonly headerFooter?: HeaderFooterOptions;
  private readonly pageSetup?: PageSetupOptions;
  private readonly tabColor?: TabColorOptions;
  private readonly sheetView?: SheetViewOptions;
  private readonly pivotTableOptions: readonly PivotTableOptions[];
  private readonly tableOptions: readonly TableOptions[];

  public constructor(options: WorksheetOptions) {
    super("worksheet");
    this.rows = options.rows ?? [];
    this.columns = options.columns ?? [];
    this.mergeCells = options.mergeCells ?? [];
    this.freezePanes = options.freezePanes;
    this.protection = options.protection;
    this.protectedRanges = options.protectedRanges ?? [];
    this.autoFilter = options.autoFilter;
    this.images = options.images ?? [];
    this.chartOptions = options.charts ?? [];
    this.dataValidations = options.dataValidations ?? [];
    this.conditionalFormats = options.conditionalFormats ?? [];
    this.hyperlinks = options.hyperlinks ?? [];
    this.comments = options.comments ?? [];
    this.headerFooter = options.headerFooter;
    this.pageSetup = options.pageSetup;
    this.tabColor = options.tabColor;
    this.sheetView = options.sheetView;
    this.pivotTableOptions = options.pivotTables ?? [];
    this.tableOptions = options.tables ?? [];
  }

  public get imageOptions(): readonly WorksheetImageOptions[] {
    return this.images;
  }

  public get charts(): readonly WorksheetChartOptions[] {
    return this.chartOptions;
  }

  public get hyperlinkOptions(): readonly HyperlinkOptions[] {
    return this.hyperlinks;
  }

  public get worksheetRows(): readonly RowOptions[] {
    return this.rows;
  }

  public get commentOptions(): readonly CommentOptions[] {
    return this.comments;
  }

  public get pivotTables(): readonly PivotTableOptions[] {
    return this.pivotTableOptions;
  }

  public get tables(): readonly TableOptions[] {
    return this.tableOptions;
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
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
        ' mc:Ignorable="x14ac xr xr2 xr3"' +
        ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"' +
        ' xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"' +
        ' xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"' +
        ' xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3">',
    ];

    // Sheet properties (tabColor, outlinePr go here)
    const hasTabColor = !!this.tabColor;
    const hasOutline = this.columns.some((c) => c.outlineLevel !== undefined);
    if (hasTabColor || hasOutline) {
      const prParts: string[] = [];
      if (this.tabColor) {
        const tc = this.tabColor;
        const tcAttrs: Record<string, string | number | boolean | undefined> = {};
        if (tc.rgb) tcAttrs.rgb = tc.rgb;
        if (tc.theme !== undefined) tcAttrs.theme = tc.theme;
        if (tc.tint !== undefined) tcAttrs.tint = tc.tint;
        prParts.push(`<tabColor${attrs(tcAttrs)}/>`);
      }
      if (hasOutline) {
        prParts.push('<outlinePr summaryBelow="1" summaryRight="1"/>');
      }
      p.push(`<sheetPr>${prParts.join("")}</sheetPr>`);
    }

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
      const svAttrs = this.buildSheetViewAttrs();
      p.push(
        `<sheetViews><sheetView${svAttrs}>`,
        `<pane ySplit="${ySplit}" xSplit="${xSplit}" topLeftCell="${topLeftCell}" activePane="${activePane}" state="frozen"/>`,
        "</sheetView></sheetViews>",
      );
    } else {
      const svAttrs = this.buildSheetViewAttrs();
      p.push(`<sheetViews><sheetView${svAttrs}/></sheetViews>`);
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
        if (col.outlineLevel !== undefined) {
          colAttrs.outlineLevel = col.outlineLevel;
        }
        if (col.collapsed) {
          colAttrs.collapsed = 1;
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

    // Sheet protection (after sheetData, before protectedRanges per XSD sequence)
    if (this.protection) {
      const prot = this.protection;
      const protAttrs: Record<string, string | number | boolean | undefined> = {};
      if (prot.password) protAttrs.password = this.hashPassword(prot.password);
      if (prot.algorithmName) protAttrs.algorithmName = prot.algorithmName;
      if (prot.hashValue) protAttrs.hashValue = prot.hashValue;
      if (prot.saltValue) protAttrs.saltValue = prot.saltValue;
      if (prot.spinCount !== undefined) protAttrs.spinCount = prot.spinCount;
      if (prot.sheet) protAttrs.sheet = 1;
      if (prot.objects) protAttrs.objects = 1;
      if (prot.scenarios) protAttrs.scenarios = 1;
      if (prot.formatCells === false) protAttrs.formatCells = 0;
      if (prot.formatColumns === false) protAttrs.formatColumns = 0;
      if (prot.formatRows === false) protAttrs.formatRows = 0;
      if (prot.insertColumns === false) protAttrs.insertColumns = 0;
      if (prot.insertRows === false) protAttrs.insertRows = 0;
      if (prot.insertHyperlinks === false) protAttrs.insertHyperlinks = 0;
      if (prot.deleteColumns === false) protAttrs.deleteColumns = 0;
      if (prot.deleteRows === false) protAttrs.deleteRows = 0;
      if (prot.selectLockedCells) protAttrs.selectLockedCells = 1;
      if (prot.sort === false) protAttrs.sort = 0;
      if (prot.autoFilter === false) protAttrs.autoFilter = 0;
      if (prot.pivotTables === false) protAttrs.pivotTables = 0;
      if (prot.selectUnlockedCells) protAttrs.selectUnlockedCells = 1;
      p.push(selfCloseElement("sheetProtection", attrs(protAttrs)));
    }

    // Protected ranges (after sheetProtection per XSD sequence)
    if (this.protectedRanges.length > 0) {
      const prParts: string[] = ["<protectedRanges>"];
      for (const pr of this.protectedRanges) {
        const prAttrs: Record<string, string | number | boolean | undefined> = {
          name: pr.name,
          sqref: pr.sqref,
        };
        if (pr.password) prAttrs.password = this.hashPassword(pr.password);
        if (pr.algorithmName) prAttrs.algorithmName = pr.algorithmName;
        if (pr.hashValue) prAttrs.hashValue = pr.hashValue;
        if (pr.saltValue) prAttrs.saltValue = pr.saltValue;
        if (pr.spinCount !== undefined) prAttrs.spinCount = pr.spinCount;
        const hasSecurityDescriptor = !!pr.securityDescriptor;
        if (hasSecurityDescriptor) {
          prParts.push(
            `<protectedRange${attrs(prAttrs)}><securityDescriptor>${escapeXml(pr.securityDescriptor!)}</securityDescriptor></protectedRange>`,
          );
        } else {
          prParts.push(selfCloseElement("protectedRange", attrs(prAttrs)));
        }
      }
      prParts.push("</protectedRanges>");
      p.push(prParts.join(""));
    }

    // Auto filter
    if (this.autoFilter) {
      if (typeof this.autoFilter === "string") {
        p.push(selfCloseElement("autoFilter", attrs({ ref: this.autoFilter })));
      } else {
        const af = this.autoFilter;
        const inner: string[] = [];
        for (const t10 of af.top10 ?? []) {
          const t10Attrs: Record<string, string | number | boolean | undefined> = { val: t10.val };
          if (t10.top === false) t10Attrs.top = 0;
          if (t10.percent) t10Attrs.percent = 1;
          inner.push(
            `<filterColumn colId="${t10.colId}"><top10${attrs(t10Attrs)}/></filterColumn>`,
          );
        }
        for (const cf of af.customFilters ?? []) {
          const cfAttrs: Record<string, string | number | boolean | undefined> = {};
          if (cf.and) cfAttrs.and = 1;
          const filters: string[] = [];
          if (cf.val !== undefined) {
            const fAttrs: Record<string, string | number | boolean | undefined> = { val: cf.val };
            if (cf.operator) fAttrs.operator = cf.operator;
            filters.push(selfCloseElement("customFilter", attrs(fAttrs)));
          }
          if (cf.val2 !== undefined) {
            filters.push(selfCloseElement("customFilter", attrs({ val: cf.val2 })));
          }
          if (filters.length > 0) {
            inner.push(
              `<filterColumn colId="${cf.colId}"><customFilters${attrs(cfAttrs)}>${filters.join("")}</customFilters></filterColumn>`,
            );
          }
        }
        if (af.sort && af.sort.length > 0) {
          const sortParts: string[] = [];
          for (const sc of af.sort) {
            const scAttrs: Record<string, string | number | boolean | undefined> = { ref: sc.ref };
            if (sc.descending) scAttrs.descending = 1;
            sortParts.push(selfCloseElement("sortCondition", attrs(scAttrs)));
          }
          inner.push(`<sortState ref="${af.ref}">${sortParts.join("")}</sortState>`);
        }
        if (inner.length > 0) {
          p.push(`<autoFilter ref="${af.ref}">`, ...inner, "</autoFilter>");
        } else {
          p.push(selfCloseElement("autoFilter", attrs({ ref: af.ref })));
        }
      }
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

          // Color scale
          if (rule.type === "colorScale" && rule.colorScale) {
            const cs = rule.colorScale;
            const inner: string[] = [];
            for (const v of cs.cfvo) {
              inner.push(this.buildCfvoXml(v));
            }
            for (const c of cs.colors) {
              inner.push(`<color rgb="FF${c}"/>`);
            }
            p.push(
              `<cfRule${attrs(ruleAttrs)}><colorScale>${inner.join("")}</colorScale></cfRule>`,
            );
          }
          // Data bar
          else if (rule.type === "dataBar" && rule.dataBar) {
            const db = rule.dataBar;
            const inner: string[] = [];
            for (const v of db.cfvo) {
              inner.push(this.buildCfvoXml(v));
            }
            inner.push(`<color rgb="FF${db.color}"/>`);
            const dbAttrs: Record<string, string | number | boolean | undefined> = {};
            if (db.minLength !== undefined && db.minLength !== 10) dbAttrs.minLength = db.minLength;
            if (db.maxLength !== undefined && db.maxLength !== 90) dbAttrs.maxLength = db.maxLength;
            if (db.showValue === false) dbAttrs.showValue = 0;
            const attrStr = Object.keys(dbAttrs).length > 0 ? attrs(dbAttrs) : "";
            p.push(
              `<cfRule${attrs(ruleAttrs)}><dataBar${attrStr}>${inner.join("")}</dataBar></cfRule>`,
            );
          }
          // Icon set
          else if (rule.type === "iconSet" && rule.iconSet) {
            const is = rule.iconSet;
            const inner: string[] = [];
            for (const v of is.cfvo) {
              inner.push(this.buildCfvoXml(v));
            }
            const isAttrs: Record<string, string | number | boolean | undefined> = {};
            if (is.iconSet !== undefined && is.iconSet !== "3TrafficLights1")
              isAttrs.iconSet = is.iconSet;
            if (is.showValue === false) isAttrs.showValue = 0;
            if (is.percent === false) isAttrs.percent = 0;
            if (is.reverse) isAttrs.reverse = 1;
            const attrStr = Object.keys(isAttrs).length > 0 ? attrs(isAttrs) : "";
            p.push(
              `<cfRule${attrs(ruleAttrs)}><iconSet${attrStr}>${inner.join("")}</iconSet></cfRule>`,
            );
          }
          // Standard rules (cellIs, containsText, expression, top10, aboveAverage)
          else {
            if (rule.formulas && rule.formulas.length > 0) {
              const formulaParts = rule.formulas.map((f) => `<formula>${escapeXml(f)}</formula>`);
              p.push(`<cfRule${attrs(ruleAttrs)}>`, ...formulaParts, "</cfRule>");
            } else {
              p.push(selfCloseElement("cfRule", attrs(ruleAttrs)));
            }
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

    // Hyperlinks — r:id numbering must match worksheet rels order (compiler handles rels)
    if (this.hyperlinks.length > 0) {
      p.push("<hyperlinks>");
      let hlIdx = 0;
      for (const hl of this.hyperlinks) {
        const hlAttrs: Record<string, string | number | boolean | undefined> = { ref: hl.cell };
        if (hl.target.type === "external") {
          hlIdx++;
          hlAttrs["r:id"] = `rId${hlIdx}`;
        } else {
          hlAttrs.location = hl.target.location;
        }
        if (hl.tooltip) hlAttrs.tooltip = hl.tooltip;
        if (hl.display) hlAttrs.display = hl.display;
        p.push(selfCloseElement("hyperlink", attrs(hlAttrs)));
      }
      p.push("</hyperlinks>");
    }

    p.push('<pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>');

    // Page setup
    if (this.pageSetup) {
      const ps = this.pageSetup;
      const psAttrs: Record<string, string | number | boolean | undefined> = {};
      if (ps.paperSize !== undefined) psAttrs.paperSize = ps.paperSize;
      if (ps.orientation && ps.orientation !== "default") psAttrs.orientation = ps.orientation;
      if (ps.scale !== undefined) psAttrs.scale = ps.scale;
      if (ps.fitToWidth !== undefined) psAttrs.fitToWidth = ps.fitToWidth;
      if (ps.fitToHeight !== undefined) psAttrs.fitToHeight = ps.fitToHeight;
      if (ps.pageOrder && ps.pageOrder !== "downThenOver") psAttrs.pageOrder = ps.pageOrder;
      if (ps.useFirstPageNumber) psAttrs.useFirstPageNumber = 1;
      if (ps.firstPageNumber !== undefined) psAttrs.firstPageNumber = ps.firstPageNumber;
      p.push(selfCloseElement("pageSetup", attrs(psAttrs)));
    }

    // Header/footer
    if (this.headerFooter) {
      const hf = this.headerFooter;
      const hfAttrs: Record<string, string | number | boolean | undefined> = {};
      if (hf.differentOddEven) hfAttrs.differentOddEven = 1;
      if (hf.differentFirst) hfAttrs.differentFirst = 1;
      const inner: string[] = [];
      if (hf.oddHeader) inner.push(`<oddHeader>${escapeXml(hf.oddHeader)}</oddHeader>`);
      if (hf.oddFooter) inner.push(`<oddFooter>${escapeXml(hf.oddFooter)}</oddFooter>`);
      if (hf.evenHeader) inner.push(`<evenHeader>${escapeXml(hf.evenHeader)}</evenHeader>`);
      if (hf.evenFooter) inner.push(`<evenFooter>${escapeXml(hf.evenFooter)}</evenFooter>`);
      if (hf.firstHeader) inner.push(`<firstHeader>${escapeXml(hf.firstHeader)}</firstHeader>`);
      if (hf.firstFooter) inner.push(`<firstFooter>${escapeXml(hf.firstFooter)}</firstFooter>`);
      if (inner.length > 0) {
        p.push(`<headerFooter${attrs(hfAttrs)}>`, ...inner, "</headerFooter>");
      } else if (hfAttrs.differentOddEven || hfAttrs.differentFirst) {
        p.push(selfCloseElement("headerFooter", attrs(hfAttrs)));
      }
    }

    p.push("</worksheet>");
    return p.join("");
  }

  /**
   * Build a <cfvo> element string for conditional formatting.
   */
  private buildCfvoXml(cfvo: CfvoOptions): string {
    const a: Record<string, string | number | boolean | undefined> = { type: cfvo.type };
    if (cfvo.val !== undefined) a.val = cfvo.val;
    if (cfvo.gte === false) a.gte = 0;
    return `<cfvo${attrs(a)}/>`;
  }

  private buildSheetViewAttrs(): string {
    const sv = this.sheetView;
    const svMap: Record<string, string | number | boolean | undefined> = {
      workbookViewId: 0,
    };
    if (sv?.tabSelected !== undefined) svMap.tabSelected = sv.tabSelected ? 1 : 0;
    else svMap.tabSelected = 1;
    if (sv?.showGridLines === false) svMap.showGridLines = 0;
    if (sv?.showRowColHeaders === false) svMap.showRowColHeaders = 0;
    if (sv?.showZeros === false) svMap.showZeros = 0;
    if (sv?.zoomScale !== undefined) svMap.zoomScale = sv.zoomScale;
    if (sv?.rightToLeft) svMap.rightToLeft = 1;
    return attrs(svMap);
  }

  /**
   * Excel legacy password hash (16-bit, little-endian hex).
   * Matches the algorithm used by ECMA-376 Part 1, §18.2.27.
   */
  private hashPassword(password: string): string {
    let hash = 0;
    for (let i = 0; i < password.length; i++) {
      const c = password.charCodeAt(i);
      hash = ((hash >> 14) & 1) + ((hash << 1) & 0x7fff);
      hash ^= c;
      hash = hash & 0x4000 ? hash ^ 0x1 : hash;
    }
    hash = ((hash >> 14) & 1) + ((hash << 1) & 0x7fff);
    hash = ((hash >> 14) & 1) + ((hash << 1) & 0x7fff);
    hash ^= password.length;
    return hash.toString(16).toUpperCase().padStart(4, "0");
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

  private dateToSerialNumber(date: Date): number {
    const epoch = new Date(1899, 11, 30);
    const msPerDay = 86400000;
    return (date.getTime() - epoch.getTime()) / msPerDay;
  }
}
