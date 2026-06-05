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
  /** Best-fit column width (CT_Col @bestFit) */
  readonly bestFit?: boolean;
  /** Phonetic text for CJK (CT_Col @phonetic) */
  readonly phonetic?: boolean;
}

export interface RowOptions {
  readonly cells?: readonly CellOptions[];
  readonly height?: number;
  readonly hidden?: boolean;
  readonly rowNumber?: number;
  /** Spans for the row, e.g. "1:15" (CT_Row @spans) */
  readonly spans?: string;
  /** Custom format applied (CT_Row @customFormat) */
  readonly customFormat?: boolean;
  /** Thick top border (CT_Row @thickTop) */
  readonly thickTop?: boolean;
  /** Thick bottom border (CT_Row @thickBot) */
  readonly thickBot?: boolean;
  /** Phonetic text (CT_Row @ph) */
  readonly ph?: boolean;
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
  /** Always calculate array (CT_CellFormula @aca) */
  readonly aca?: boolean;
  /** 2-D data table (CT_CellFormula @dt2D) */
  readonly dt2D?: boolean;
  /** Data table row (CT_CellFormula @dtr) */
  readonly dtr?: boolean;
  /** Delete input cell 1 (CT_CellFormula @del1) */
  readonly del1?: boolean;
  /** Delete input cell 2 (CT_CellFormula @del2) */
  readonly del2?: boolean;
  /** Input cell 1 reference (CT_CellFormula @r1) */
  readonly r1?: string;
  /** Input cell 2 reference (CT_CellFormula @r2) */
  readonly r2?: string;
  /** Calculate cell (CT_CellFormula @ca) */
  readonly ca?: boolean;
  /** Array formula context (CT_CellFormula @bx) */
  readonly bx?: boolean;
}

/** Input cell for a what-if scenario (maps to CT_InputCells). */
export interface ScenarioCellOptions {
  /** Cell reference, e.g. "B2" */
  readonly r: string;
  /** Cell value for this scenario */
  readonly val: string | number;
  /** Whether the value is deleted */
  readonly deleted?: boolean;
}

/** A single what-if scenario (maps to CT_Scenario). */
export interface ScenarioDefinition {
  /** Scenario name */
  readonly name: string;
  /** Input cells with their values for this scenario */
  readonly inputCells: readonly ScenarioCellOptions[];
  /** Sort/order count */
  readonly count?: number;
  /** Creator user name */
  readonly user?: string;
  /** Comment */
  readonly comment?: string;
  /** Whether the scenario is hidden */
  readonly hidden?: boolean;
  /** Whether the scenario is locked */
  readonly locked?: boolean;
}

/** Scenarios for what-if analysis (maps to CT_Scenarios). */
export interface ScenarioOptions {
  /** Named scenarios */
  readonly scenarios: readonly ScenarioDefinition[];
  /** Current scenario index (0-based) */
  readonly current?: number;
  /** Show scenario index (0-based) */
  readonly show?: number;
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
  /** Window protection (CT_SheetView @windowProtection) */
  readonly windowProtection?: boolean;
  /** Show formulas instead of values (CT_SheetView @showFormulas) */
  readonly showFormulas?: boolean;
  /** Show ruler (CT_SheetView @showRuler) */
  readonly showRuler?: boolean;
  /** Show outline symbols (CT_SheetView @showOutlineSymbols) */
  readonly showOutlineSymbols?: boolean;
  /** Default grid color (CT_SheetView @defaultGridColor) */
  readonly defaultGridColor?: boolean;
  /** Show white space (CT_SheetView @showWhiteSpace) */
  readonly showWhiteSpace?: boolean;
  /** View type (CT_SheetView @view) */
  readonly view?: "normal" | "pageBreakPreview" | "pageLayout";
  /** Tab color ID (CT_SheetView @colorId) */
  readonly colorId?: number;
  /** Zoom scale for normal view (CT_SheetView @zoomScaleNormal) */
  readonly zoomScaleNormal?: number;
  /** Zoom scale for sheet layout view (CT_SheetView @zoomScaleSheetLayoutView) */
  readonly zoomScaleSheetLayoutView?: number;
  /** Zoom scale for page layout view (CT_SheetView @zoomScalePageLayoutView) */
  readonly zoomScalePageLayoutView?: number;
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
  /** Scale header/footer with document (CT_HeaderFooter @scaleWithDoc) */
  readonly scaleWithDoc?: boolean;
  /** Align with page margins (CT_HeaderFooter @alignWithMargins) */
  readonly alignWithMargins?: boolean;
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
  /** Paper height (CT_PageSetup @paperHeight) */
  readonly paperHeight?: number;
  /** Paper width (CT_PageSetup @paperWidth) */
  readonly paperWidth?: number;
  /** Use printer defaults (CT_PageSetup @usePrinterDefaults) */
  readonly usePrinterDefaults?: boolean;
  /** Black and white printing (CT_PageSetup @blackAndWhite) */
  readonly blackAndWhite?: boolean;
  /** Draft quality printing (CT_PageSetup @draft) */
  readonly draft?: boolean;
  /** Print cell comments mode (CT_PageSetup @cellComments) */
  readonly cellComments?: "none" | "asDisplayed" | "atEnd";
  /** Print error display mode (CT_PageSetup @errors) */
  readonly errors?: "displayed" | "blank" | "dash" | "NA";
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
  /** Error style (CT_DataValidation @errorStyle) */
  readonly errorStyle?: "stop" | "warning" | "information";
  /** IME mode (CT_DataValidation @imeMode) */
  readonly imeMode?:
    | "noControl"
    | "on"
    | "off"
    | "disabled"
    | "hiragana"
    | "fullKatakana"
    | "halfKatakana"
    | "fullAlpha"
    | "halfAlpha"
    | "fullHangul"
    | "halfHangul";
  /** Show drop-down (CT_DataValidation @showDropDown — note inverted semantics in OOXML) */
  readonly showDropDown?: boolean;
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
  /** Stop if true — skip remaining rules (CT_CfRule @stopIfTrue) */
  readonly stopIfTrue?: boolean;
  /** Time period for date-based highlighting (CT_CfRule @timePeriod) */
  readonly timePeriod?:
    | "today"
    | "yesterday"
    | "tomorrow"
    | "last7Days"
    | "thisMonth"
    | "lastMonth"
    | "nextMonth"
    | "thisWeek"
    | "lastWeek"
    | "nextWeek";
  /** Rank for top/bottom rules (CT_CfRule @rank) */
  readonly rank?: number;
  /** Equal average flag (CT_CfRule @equalAverage) */
  readonly equalAverage?: boolean;
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
  /** Sort by (CT_SortCondition @sortBy) */
  readonly sortBy?: "value" | "cellColor" | "fontColor" | "icon";
  /** Custom sort list (CT_SortCondition @customList) */
  readonly customList?: string;
  /** Icon set index (CT_SortCondition @iconId) */
  readonly iconId?: number;
}

export interface AutoFilterOptions {
  /** Range, e.g. "A1:D10" */
  readonly ref: string;
  readonly top10?: readonly Top10FilterOptions[];
  readonly customFilters?: readonly CustomFilterOptions[];
  readonly sort?: readonly SortCondition[];
  /** Sort state options */
  readonly sortState?: SortStateOptions;
}

/** Sort state configuration (CT_SortState) */
export interface SortStateOptions {
  /** Column sort mode (CT_SortState @columnSort) */
  readonly columnSort?: boolean;
  /** Case sensitive sorting (CT_SortState @caseSensitive) */
  readonly caseSensitive?: boolean;
  /** Sort method (CT_SortState @sortMethod) */
  readonly sortMethod?: "pinYin" | "stroke";
}

/** Print options (CT_PrintOptions) */
export interface PrintOptions {
  /** Center horizontally on page */
  readonly horizontalCentered?: boolean;
  /** Center vertically on page */
  readonly verticalCentered?: boolean;
  /** Print row/column headings */
  readonly headings?: boolean;
  /** Print grid lines */
  readonly gridLines?: boolean;
  /** Grid lines set flag */
  readonly gridLinesSet?: boolean;
}

/** Sheet format properties (CT_SheetFormatPr) */
export interface SheetFormatPrOptions {
  /** Base column width (CT_SheetFormatPr @baseColWidth) */
  readonly baseColWidth?: number;
  /** Default column width (CT_SheetFormatPr @defaultColWidth) */
  readonly defaultColWidth?: number;
  /** Default row height */
  readonly defaultRowHeight?: number;
  /** Zero height rows hidden (CT_SheetFormatPr @zeroHeight) */
  readonly zeroHeight?: boolean;
  /** Thick top borders (CT_SheetFormatPr @thickTop) */
  readonly thickTop?: boolean;
  /** Thick bottom borders (CT_SheetFormatPr @thickBottom) */
  readonly thickBottom?: boolean;
  /** Outline level row (CT_SheetFormatPr @outlineLevelRow) */
  readonly outlineLevelRow?: number;
  /** Outline level column (CT_SheetFormatPr @outlineLevelCol) */
  readonly outlineLevelCol?: number;
}

/** Sheet properties extended options (CT_SheetPr attributes) */
export interface SheetPrOptions {
  /** Sync horizontal scroll (CT_SheetPr @syncHorizontal) */
  readonly syncHorizontal?: boolean;
  /** Sync vertical scroll (CT_SheetPr @syncVertical) */
  readonly syncVertical?: boolean;
  /** Sync reference (CT_SheetPr @syncRef) */
  readonly syncRef?: string;
  /** Transition evaluation mode (CT_SheetPr @transitionEvaluation) */
  readonly transitionEvaluation?: boolean;
  /** Transition entry mode (CT_SheetPr @transitionEntry) */
  readonly transitionEntry?: boolean;
  /** Published to server (CT_SheetPr @published) */
  readonly published?: boolean;
  /** Filter mode (CT_SheetPr @filterMode) */
  readonly filterMode?: boolean;
  /** Enable format conditions calculation (CT_SheetPr @enableFormatConditionsCalculation) */
  readonly enableFormatConditionsCalculation?: boolean;
  /** Outline apply styles (CT_OutlinePr @applyStyles) */
  readonly outlineApplyStyles?: boolean;
  /** Outline show symbols (CT_OutlinePr @showOutlineSymbols) */
  readonly outlineShowSymbols?: boolean;
}

/** An ignored error entry — suppresses specific Excel error checks for a range. */
export interface IgnoredErrorOptions {
  /** Cell range, e.g. "A1:A10" (required) */
  readonly sqref: string;
  readonly evalError?: boolean;
  readonly twoDigitTextYear?: boolean;
  readonly numberStoredAsText?: boolean;
  readonly formula?: boolean;
  readonly formulaRange?: boolean;
  readonly unlockedFormula?: boolean;
  readonly emptyCellReference?: boolean;
  readonly listDataValidation?: boolean;
  readonly calculatedColumn?: boolean;
}

/** Phonetic properties for CJK text (CT_PhoneticPr) */
export interface PhoneticPrOptions {
  /** Font ID from the styles table (required) */
  readonly fontId: number;
  /** Phonetic type (default: "fullwidthKatakana") */
  readonly type?: "fullwidthKatakana" | "halfwidthKatakana" | "Hiragana" | "noConversion";
  /** Alignment (default: "left") */
  readonly alignment?: "left" | "center" | "distributed";
}

/** Background image for a worksheet */
export interface SheetBackgroundImageOptions {
  readonly data: Uint8Array;
  readonly type: "png" | "jpeg" | "jpg";
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
  /** What-if scenarios */
  readonly scenarios?: ScenarioOptions;
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
  /** Ignored errors — suppress specific Excel error checks for cell ranges */
  readonly ignoredErrors?: readonly IgnoredErrorOptions[];
  /** Phonetic properties for CJK text */
  readonly phoneticPr?: PhoneticPrOptions;
  /** Background image for the worksheet */
  readonly backgroundImage?: SheetBackgroundImageOptions;
  /** Print options (CT_PrintOptions) */
  readonly printOptions?: PrintOptions;
  /** Sheet format properties (CT_SheetFormatPr) */
  readonly sheetFormatPr?: SheetFormatPrOptions;
  /** Sheet extended properties (CT_SheetPr attributes) */
  readonly sheetPr?: SheetPrOptions;
}

export class Worksheet extends IgnoreIfEmptyXmlComponent {
  private readonly rows: readonly RowOptions[];
  private readonly columns: readonly ColumnOptions[];
  private readonly mergeCells: readonly MergeCellOptions[];
  private readonly freezePanes?: FreezePaneOptions;
  private readonly protection?: SheetProtectionOptions;
  private readonly protectedRanges: readonly ProtectedRangeOptions[];
  private readonly scenarioOpts?: ScenarioOptions;
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
  private readonly ignoredErrors: readonly IgnoredErrorOptions[];
  private readonly phoneticPr?: PhoneticPrOptions;
  private readonly backgroundImage?: SheetBackgroundImageOptions;
  private readonly printOptions?: PrintOptions;
  private readonly sheetFormatPr?: SheetFormatPrOptions;
  private readonly sheetPr?: SheetPrOptions;

  public constructor(options: WorksheetOptions) {
    super("worksheet");
    this.rows = options.rows ?? [];
    this.columns = options.columns ?? [];
    this.mergeCells = options.mergeCells ?? [];
    this.freezePanes = options.freezePanes;
    this.protection = options.protection;
    this.protectedRanges = options.protectedRanges ?? [];
    this.scenarioOpts = options.scenarios;
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
    this.ignoredErrors = options.ignoredErrors ?? [];
    this.phoneticPr = options.phoneticPr;
    this.backgroundImage = options.backgroundImage;
    this.printOptions = options.printOptions;
    this.sheetFormatPr = options.sheetFormatPr;
    this.sheetPr = options.sheetPr;
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

  public get background(): SheetBackgroundImageOptions | undefined {
    return this.backgroundImage;
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
    const sp = this.sheetPr;
    const hasSheetPrAttrs =
      sp &&
      (sp.syncHorizontal ||
        sp.syncVertical ||
        sp.syncRef ||
        sp.transitionEvaluation ||
        sp.transitionEntry ||
        sp.published ||
        sp.filterMode ||
        sp.enableFormatConditionsCalculation);
    if (hasTabColor || hasOutline || hasSheetPrAttrs) {
      const prParts: string[] = [];
      const prAttrs: Record<string, string | number | boolean | undefined> = {};
      if (sp?.syncHorizontal) prAttrs.syncHorizontal = 1;
      if (sp?.syncVertical) prAttrs.syncVertical = 1;
      if (sp?.syncRef) prAttrs.syncRef = sp.syncRef;
      if (sp?.transitionEvaluation) prAttrs.transitionEvaluation = 1;
      if (sp?.transitionEntry) prAttrs.transitionEntry = 1;
      if (sp?.published) prAttrs.published = 1;
      if (sp?.filterMode) prAttrs.filterMode = 1;
      if (sp?.enableFormatConditionsCalculation) prAttrs.enableFormatConditionsCalculation = 1;
      if (this.tabColor) {
        const tc = this.tabColor;
        const tcAttrs: Record<string, string | number | boolean | undefined> = {};
        if (tc.rgb) tcAttrs.rgb = tc.rgb;
        if (tc.theme !== undefined) tcAttrs.theme = tc.theme;
        if (tc.tint !== undefined) tcAttrs.tint = tc.tint;
        prParts.push(`<tabColor${attrs(tcAttrs)}/>`);
      }
      if (hasOutline) {
        const outAttrs: Record<string, string | number | boolean | undefined> = {
          summaryBelow: 1,
          summaryRight: 1,
        };
        if (sp?.outlineApplyStyles) outAttrs.applyStyles = 1;
        if (sp?.outlineShowSymbols === false) outAttrs.showOutlineSymbols = 0;
        prParts.push(`<outlinePr${attrs(outAttrs)}/>`);
      }
      const prAttrStr = Object.keys(prAttrs).length > 0 ? attrs(prAttrs) : "";
      p.push(`<sheetPr${prAttrStr}>${prParts.join("")}</sheetPr>`);
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
    if (this.sheetFormatPr) {
      const sfp = this.sheetFormatPr;
      const sfpAttrs: Record<string, string | number | boolean | undefined> = {};
      if (sfp.baseColWidth !== undefined) sfpAttrs.baseColWidth = sfp.baseColWidth;
      if (sfp.defaultColWidth !== undefined) sfpAttrs.defaultColWidth = sfp.defaultColWidth;
      sfpAttrs.defaultRowHeight = sfp.defaultRowHeight ?? 15;
      if (sfp.zeroHeight) sfpAttrs.zeroHeight = 1;
      if (sfp.thickTop) sfpAttrs.thickTop = 1;
      if (sfp.thickBottom) sfpAttrs.thickBottom = 1;
      if (sfp.outlineLevelRow !== undefined) sfpAttrs.outlineLevelRow = sfp.outlineLevelRow;
      if (sfp.outlineLevelCol !== undefined) sfpAttrs.outlineLevelCol = sfp.outlineLevelCol;
      p.push(`<sheetFormatPr${attrs(sfpAttrs)}/>`);
    } else {
      p.push('<sheetFormatPr defaultRowHeight="15"/>');
    }

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
        if (col.bestFit) {
          colAttrs.bestFit = 1;
        }
        if (col.phonetic) {
          colAttrs.phonetic = 1;
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
      if (rowOpts.spans) rowAttrs.spans = rowOpts.spans;
      if (rowOpts.customFormat) rowAttrs.customFormat = 1;
      if (rowOpts.thickTop) rowAttrs.thickTop = 1;
      if (rowOpts.thickBot) rowAttrs.thickBot = 1;
      if (rowOpts.ph) rowAttrs.ph = 1;

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

    // Scenarios (what-if analysis)
    if (this.scenarioOpts) {
      const scParts: string[] = ["<scenarios"];
      const scAttrs: Record<string, string | number> = {};
      if (this.scenarioOpts.current !== undefined) scAttrs.current = this.scenarioOpts.current;
      if (this.scenarioOpts.show !== undefined) scAttrs.show = this.scenarioOpts.show;
      scParts[0] = `<scenarios${attrs(scAttrs)}>`;

      for (const scenario of this.scenarioOpts.scenarios) {
        const sAttrs: Record<string, string | number | boolean | undefined> = {
          name: scenario.name,
        };
        if (scenario.count !== undefined) sAttrs.count = scenario.count;
        if (scenario.user) sAttrs.user = scenario.user;
        if (scenario.comment) sAttrs.comment = scenario.comment;
        if (scenario.hidden) sAttrs.hidden = true;
        if (scenario.locked) sAttrs.locked = true;

        const sParts: string[] = [`<scenario${attrs(sAttrs)}>`];
        for (const cell of scenario.inputCells) {
          const icAttrs: Record<string, string | number | boolean | undefined> = {
            r: cell.r,
            val: String(cell.val),
          };
          if (cell.deleted) icAttrs.deleted = true;
          sParts.push(`<inputCells${attrs(icAttrs)}/>`);
        }
        sParts.push("</scenario>");
        scParts.push(sParts.join(""));
      }
      scParts.push("</scenarios>");
      p.push(scParts.join(""));
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

    // Phonetic properties (after mergeCells per XSD sequence)
    if (this.phoneticPr) {
      const pp = this.phoneticPr;
      const ppAttrs: Record<string, string | number> = { fontId: pp.fontId };
      if (pp.type && pp.type !== "fullwidthKatakana") ppAttrs.type = pp.type;
      if (pp.alignment && pp.alignment !== "left") ppAttrs.alignment = pp.alignment;
      p.push(selfCloseElement("phoneticPr", attrs(ppAttrs)));
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
          if (rule.stopIfTrue) ruleAttrs.stopIfTrue = 1;
          if (rule.timePeriod) ruleAttrs.timePeriod = rule.timePeriod;
          if (rule.rank !== undefined) ruleAttrs.rank = rule.rank;
          if (rule.equalAverage) ruleAttrs.equalAverage = 1;

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
        if (dv.errorStyle) dvAttrs.errorStyle = dv.errorStyle;
        if (dv.imeMode) dvAttrs.imeMode = dv.imeMode;
        if (dv.showDropDown) dvAttrs.showDropDown = 1;
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

    // Print options
    if (this.printOptions) {
      const po = this.printOptions;
      const poAttrs: Record<string, string | number | boolean | undefined> = {};
      if (po.horizontalCentered) poAttrs.horizontalCentered = 1;
      if (po.verticalCentered) poAttrs.verticalCentered = 1;
      if (po.headings) poAttrs.headings = 1;
      if (po.gridLines) poAttrs.gridLines = 1;
      if (po.gridLinesSet === false) poAttrs.gridLinesSet = 0;
      p.push(selfCloseElement("printOptions", attrs(poAttrs)));
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
      if (ps.paperHeight !== undefined) psAttrs.paperHeight = ps.paperHeight;
      if (ps.paperWidth !== undefined) psAttrs.paperWidth = ps.paperWidth;
      if (ps.usePrinterDefaults) psAttrs.usePrinterDefaults = 1;
      if (ps.blackAndWhite) psAttrs.blackAndWhite = 1;
      if (ps.draft) psAttrs.draft = 1;
      if (ps.cellComments && ps.cellComments !== "none") psAttrs.cellComments = ps.cellComments;
      if (ps.errors && ps.errors !== "displayed") psAttrs.errors = ps.errors;
      p.push(selfCloseElement("pageSetup", attrs(psAttrs)));
    }

    // Header/footer
    if (this.headerFooter) {
      const hf = this.headerFooter;
      const hfAttrs: Record<string, string | number | boolean | undefined> = {};
      if (hf.differentOddEven) hfAttrs.differentOddEven = 1;
      if (hf.differentFirst) hfAttrs.differentFirst = 1;
      if (hf.scaleWithDoc === false) hfAttrs.scaleWithDoc = 0;
      if (hf.alignWithMargins === false) hfAttrs.alignWithMargins = 0;
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

    // Ignored errors (after headerFooter per XSD sequence)
    if (this.ignoredErrors.length > 0) {
      const ieParts: string[] = ["<ignoredErrors>"];
      for (const ie of this.ignoredErrors) {
        const ieAttrs: Record<string, string | number | boolean | undefined> = {
          sqref: ie.sqref,
        };
        if (ie.evalError) ieAttrs.evalError = 1;
        if (ie.twoDigitTextYear) ieAttrs.twoDigitTextYear = 1;
        if (ie.numberStoredAsText) ieAttrs.numberStoredAsText = 1;
        if (ie.formula) ieAttrs.formula = 1;
        if (ie.formulaRange) ieAttrs.formulaRange = 1;
        if (ie.unlockedFormula) ieAttrs.unlockedFormula = 1;
        if (ie.emptyCellReference) ieAttrs.emptyCellReference = 1;
        if (ie.listDataValidation) ieAttrs.listDataValidation = 1;
        if (ie.calculatedColumn) ieAttrs.calculatedColumn = 1;
        ieParts.push(selfCloseElement("ignoredError", attrs(ieAttrs)));
      }
      ieParts.push("</ignoredErrors>");
      p.push(ieParts.join(""));
    }

    // Background picture placeholder — compiler replaces with <picture r:id="rIdN"/>
    if (this.backgroundImage) {
      p.push("<!--BACKGROUND_PICTURE-->");
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
    if (sv?.windowProtection) svMap.windowProtection = 1;
    if (sv?.showFormulas) svMap.showFormulas = 1;
    if (sv?.showRuler === false) svMap.showRuler = 0;
    if (sv?.showOutlineSymbols === false) svMap.showOutlineSymbols = 0;
    if (sv?.defaultGridColor === false) svMap.defaultGridColor = 0;
    if (sv?.showWhiteSpace === false) svMap.showWhiteSpace = 0;
    if (sv?.view) svMap.view = sv.view;
    if (sv?.colorId !== undefined) svMap.colorId = sv.colorId;
    if (sv?.zoomScaleNormal !== undefined) svMap.zoomScaleNormal = sv.zoomScaleNormal;
    if (sv?.zoomScaleSheetLayoutView !== undefined)
      svMap.zoomScaleSheetLayoutView = sv.zoomScaleSheetLayoutView;
    if (sv?.zoomScalePageLayoutView !== undefined)
      svMap.zoomScalePageLayoutView = sv.zoomScalePageLayoutView;
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
    if (fOpts.aca) fAttrs.aca = 1;
    if (fOpts.dt2D) fAttrs.dt2D = 1;
    if (fOpts.dtr) fAttrs.dtr = 1;
    if (fOpts.del1) fAttrs.del1 = 1;
    if (fOpts.del2) fAttrs.del2 = 1;
    if (fOpts.r1) fAttrs.r1 = fOpts.r1;
    if (fOpts.r2) fAttrs.r2 = fOpts.r2;
    if (fOpts.ca) fAttrs.ca = 1;
    if (fOpts.bx) fAttrs.bx = 1;

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
