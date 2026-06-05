import type { SharedStrings } from "@file/shared-strings";
import { buildRstXml } from "@file/shared-strings";
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

/** Rich text run properties (CT_RPrElt). */
export interface RichTextRunPrOptions {
  /** Font name (CT_FontName → rFont) */
  readonly font?: string;
  /** Character set (CT_IntProperty) */
  readonly charset?: number;
  /** Font family (CT_IntProperty) */
  readonly family?: number;
  /** Bold */
  readonly bold?: boolean;
  /** Italic */
  readonly italic?: boolean;
  /** Strikethrough */
  readonly strike?: boolean;
  /** Outline */
  readonly outline?: boolean;
  /** Shadow */
  readonly shadow?: boolean;
  /** Condense */
  readonly condense?: boolean;
  /** Extend */
  readonly extend?: boolean;
  /** Font color (hex RGB, e.g. "FF0000") */
  readonly color?: string;
  /** Font size in points */
  readonly size?: number;
  /** Underline type */
  readonly underline?: "single" | "double" | "singleAccounting" | "doubleAccounting" | "none";
  /** Vertical alignment */
  readonly vertAlign?: "superscript" | "subscript" | "baseline";
  /** Font scheme */
  readonly scheme?: "major" | "minor" | "none";
}

/** A single rich text run (CT_RElt). */
export interface RichTextRunOptions {
  /** Run properties (optional = inherits from parent) */
  readonly properties?: RichTextRunPrOptions;
  /** Run text content */
  readonly text: string;
}

/** Phonetics run for CJK (CT_PhoneticRun → rPh). */
export interface PhoneticRunOptions {
  /** Start byte offset in base text */
  readonly sb: number;
  /** End byte offset in base text */
  readonly eb: number;
  /** Phonetic text */
  readonly text: string;
}

/** Rich text content (CT_Rst). Either plain text or rich runs. */
export interface RichTextOptions {
  /** Plain text (mutually exclusive with runs) */
  readonly text?: string;
  /** Rich text runs (mutually exclusive with text) */
  readonly runs?: readonly RichTextRunOptions[];
  /** Phonetic runs for CJK (CT_PhoneticRun) */
  readonly phonetics?: readonly PhoneticRunOptions[];
}

export interface CellOptions {
  readonly value?: string | number | boolean | Date | RichTextOptions | null;
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

/** Object anchor (CT_ObjectAnchor). */
export interface ObjectAnchorOptions {
  /** Move with cells (default: false) */
  readonly moveWithCells?: boolean;
  /** Size with cells (default: false) */
  readonly sizeWithCells?: boolean;
}

/** Comment property (CT_CommentPr). */
export interface CommentPrOptions {
  /** Locked */
  readonly locked?: boolean;
  /** Default size */
  readonly defaultSize?: boolean;
  /** Print */
  readonly print?: boolean;
  /** Disabled */
  readonly disabled?: boolean;
  /** Auto fill */
  readonly autoFill?: boolean;
  /** Auto line */
  readonly autoLine?: boolean;
  /** Alt text */
  readonly altText?: string;
  /** Text horizontal alignment */
  readonly textHAlign?: "left" | "center" | "right" | "justify" | "distributed";
  /** Text vertical alignment */
  readonly textVAlign?: "top" | "center" | "bottom" | "justify" | "distributed";
  /** Lock text */
  readonly lockText?: boolean;
  /** Justify last line */
  readonly justLastX?: boolean;
  /** Auto scale */
  readonly autoScale?: boolean;
  /** Object anchor position */
  readonly anchor?: ObjectAnchorOptions;
}

export interface CommentOptions {
  /** Cell reference, e.g. "A1" */
  readonly cell: string;
  /** Author name */
  readonly author: string;
  /** Comment text (plain string or rich text) */
  readonly text: string | RichTextOptions;
  /** Comment properties (CT_CommentPr) */
  readonly commentPr?: CommentPrOptions;
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
  /** Color filters (CT_ColorFilter) */
  readonly colorFilters?: readonly ColorFilterOptions[];
  /** Icon filters (CT_IconFilter) */
  readonly iconFilters?: readonly IconFilterOptions[];
  /** Dynamic filters (CT_DynamicFilter) */
  readonly dynamicFilters?: readonly DynamicFilterOptions[];
  /** Date group items in filters (CT_DateGroupItem) */
  readonly dateGroupItems?: readonly DateGroupFilterOptions[];
}

/** Color filter (CT_ColorFilter) */
export interface ColorFilterOptions {
  /** Column ID */
  readonly colId: number;
  /** Cell color RGB (dxfId used if not set) */
  readonly dxfId?: number;
  /** Filter by cell color (CT_ColorFilter @cellColor) */
  readonly cellColor?: boolean;
}

/** Icon filter (CT_IconFilter) */
export interface IconFilterOptions {
  /** Column ID */
  readonly colId: number;
  /** Icon set index (CT_IconFilter @iconSet) */
  readonly iconSet: number;
  /** Icon ID within set (CT_IconFilter @iconId) */
  readonly iconId?: number;
}

/** Dynamic filter (CT_DynamicFilter) */
export interface DynamicFilterOptions {
  /** Column ID */
  readonly colId: number;
  /** Dynamic filter type (CT_DynamicFilter @type) */
  readonly type:
    | "null"
    | "aboveAverage"
    | "belowAverage"
    | "tomorrow"
    | "today"
    | "yesterday"
    | "nextWeek"
    | "thisWeek"
    | "lastWeek"
    | "nextMonth"
    | "thisMonth"
    | "lastMonth"
    | "nextQuarter"
    | "thisQuarter"
    | "lastQuarter"
    | "nextYear"
    | "thisYear"
    | "lastYear"
    | "yearToDate"
    | "Q1"
    | "Q2"
    | "Q3"
    | "Q4"
    | "M1"
    | "M2"
    | "M3"
    | "M4"
    | "M5"
    | "M6"
    | "M7"
    | "M8"
    | "M9"
    | "M10"
    | "M11"
    | "M12";
  /** Max value (CT_DynamicFilter @val) */
  readonly val?: number;
  /** Max value as date ISO string (CT_DynamicFilter @maxVal) */
  readonly maxVal?: number;
}

/** Date group filter item (CT_DateGroupItem) */
export interface DateGroupFilterOptions {
  /** Column ID */
  readonly colId: number;
  /** Date grouping level (CT_DateGroupItem @dateTimeGrouping) */
  readonly dateTimeGrouping: "year" | "month" | "day" | "hour" | "minute" | "second";
  /** Year (CT_DateGroupItem @year) */
  readonly year?: number;
  /** Month (1-12, CT_DateGroupItem @month) */
  readonly month?: number;
  /** Day (1-31, CT_DateGroupItem @day) */
  readonly day?: number;
  /** Hour (0-23, CT_DateGroupItem @hour) */
  readonly hour?: number;
  /** Minute (0-59, CT_DateGroupItem @minute) */
  readonly minute?: number;
  /** Second (0-59, CT_DateGroupItem @second) */
  readonly second?: number;
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

/** Page break entry (CT_Break) */
export interface PageBreakOptions {
  /** Row or column ID (1-based) */
  readonly id: number;
  /** Min value (CT_Break @min) */
  readonly min?: number;
  /** Max value (CT_Break @max) */
  readonly max?: number;
  /** Manual break (CT_Break @man) */
  readonly manual?: boolean;
  /** Pivot break (CT_Break @pt) */
  readonly pivot?: boolean;
}

/** Selection in sheet view (CT_Selection) */
export interface SelectionOptions {
  /** Pane (CT_Selection @pane) */
  readonly pane?: "bottomRight" | "topRight" | "bottomLeft" | "topLeft";
  /** Active cell (CT_Selection @activeCell) */
  readonly activeCell?: string;
  /** Active cell index (CT_Selection @activeCellId) */
  readonly activeCellId?: number;
  /** Selected range (CT_Selection @sqref) */
  readonly sqref?: string;
}

/** Custom sheet view (CT_CustomSheetView) */
export interface CustomSheetViewOptions {
  /** GUID identifier (required, CT_CustomSheetView @guid) */
  readonly guid: string;
  /** Zoom scale (CT_CustomSheetView @scale) */
  readonly scale?: number;
  /** Show page breaks (CT_CustomSheetView @showPageBreaks) */
  readonly showPageBreaks?: boolean;
  /** Show formulas (CT_CustomSheetView @showFormulas) */
  readonly showFormulas?: boolean;
  /** Show grid lines (CT_CustomSheetView @showGridLines) */
  readonly showGridLines?: boolean;
  /** Show row/column headers (CT_CustomSheetView @showRowCol) */
  readonly showRowColHeaders?: boolean;
  /** Show outline symbols (CT_CustomSheetView @outlineSymbols) */
  readonly outlineSymbols?: boolean;
  /** Show zero values (CT_CustomSheetView @zeroValues) */
  readonly zeroValues?: boolean;
  /** Fit to page (CT_CustomSheetView @fitToPage) */
  readonly fitToPage?: boolean;
  /** Print area (CT_CustomSheetView @printArea) */
  readonly printArea?: boolean;
  /** Filter applied (CT_CustomSheetView @filter) */
  readonly filter?: boolean;
  /** Show auto filter (CT_CustomSheetView @showAutoFilter) */
  readonly showAutoFilter?: boolean;
  /** Hidden rows (CT_CustomSheetView @hiddenRows) */
  readonly hiddenRows?: boolean;
  /** Hidden columns (CT_CustomSheetView @hiddenColumns) */
  readonly hiddenColumns?: boolean;
  /** Sheet state (CT_CustomSheetView @state) */
  readonly state?: "visible" | "hidden" | "veryHidden";
  /** Filter unique (CT_CustomSheetView @filterUnique) */
  readonly filterUnique?: boolean;
  /** View type (CT_CustomSheetView @view) */
  readonly view?: "normal" | "pageBreakPreview" | "pageLayout";
}

/** Cell watch entry (CT_CellWatch) */
export interface CellWatchOptions {
  /** Cell reference, e.g. "A1" */
  readonly r: string;
}

/** Data consolidation (CT_DataConsolidate) */
export interface DataConsolidateOptions {
  /** Consolidation function (CT_DataConsolidate @function) */
  readonly function?:
    | "average"
    | "count"
    | "countNums"
    | "max"
    | "min"
    | "product"
    | "stdDev"
    | "stdDevp"
    | "sum"
    | "var"
    | "varp";
  /** Use top row labels (CT_DataConsolidate @startLabels) */
  readonly topLabels?: boolean;
  /** Use left column labels (CT_DataConsolidate @leftLabels) */
  readonly leftLabels?: boolean;
  /** Use labels in first row (CT_DataConsolidate @startLabels alias) */
  readonly startLabels?: boolean;
  /** Link to source data (CT_DataConsolidate @link) */
  readonly link?: boolean;
  /** Source data references */
  readonly refs?: readonly string[];
}

/** Drawing in header/footer (CT_DrawingHF) */
export interface DrawingHfOptions {
  /** Relationship ID for the drawing (required) */
  readonly rId: string;
  /** Left header odd (CT_DrawingHF @lho) */
  readonly lho?: number;
  /** Left header even (CT_DrawingHF @lhe) */
  readonly lhe?: number;
  /** Left header first (CT_DrawingHF @lhf) */
  readonly lhf?: number;
  /** Center header odd (CT_DrawingHF @cho) */
  readonly cho?: number;
  /** Center header even (CT_DrawingHF @che) */
  readonly che?: number;
  /** Center header first (CT_DrawingHF @chf) */
  readonly chf?: number;
  /** Right header odd (CT_DrawingHF @rho) */
  readonly rho?: number;
  /** Right header even (CT_DrawingHF @rhe) */
  readonly rhe?: number;
  /** Right header first (CT_DrawingHF @rhf) */
  readonly rhf?: number;
  /** Left footer odd (CT_DrawingHF @lfo) */
  readonly lfo?: number;
  /** Left footer even (CT_DrawingHF @lfe) */
  readonly lfe?: number;
  /** Left footer first (CT_DrawingHF @lff) */
  readonly lff?: number;
  /** Center footer odd (CT_DrawingHF @cfo) */
  readonly cfo?: number;
  /** Center footer even (CT_DrawingHF @cfe) */
  readonly cfe?: number;
  /** Center footer first (CT_DrawingHF @cff) */
  readonly cff?: number;
  /** Right footer odd (CT_DrawingHF @rfo) */
  readonly rfo?: number;
  /** Right footer even (CT_DrawingHF @rfe) */
  readonly rfe?: number;
  /** Right footer first (CT_DrawingHF @rff) */
  readonly rff?: number;
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
  /** Row page breaks (CT_PageBreaks) */
  readonly rowBreaks?: readonly PageBreakOptions[];
  /** Column page breaks (CT_PageBreaks) */
  readonly colBreaks?: readonly PageBreakOptions[];
  /** Custom sheet views (CT_CustomSheetViews) */
  readonly customSheetViews?: readonly CustomSheetViewOptions[];
  /** Cell watches (CT_CellWatches) */
  readonly cellWatches?: readonly CellWatchOptions[];
  /** Data consolidation (CT_DataConsolidate) */
  readonly dataConsolidate?: DataConsolidateOptions;
  /** OLE embedded range (CT_OleSize) */
  readonly oleSize?: string;
  /** Drawing in header/footer (CT_DrawingHF) */
  readonly drawingHF?: DrawingHfOptions;
  /** Legacy drawing for header/footer r:id (CT_LegacyDrawingHF) */
  readonly legacyDrawingHF?: string;
  /** Selection in sheet view (CT_Selection) */
  readonly selection?: SelectionOptions;
  /** Sheet calc properties (CT_SheetCalcPr) */
  readonly sheetCalcPr?: SheetCalcPrOptions;
  /** Extension list (extLst) */
  readonly ext?: string;
  /** Control objects (CT_Controls) */
  readonly controls?: readonly ControlOptions[];
  /** Custom sheet properties (CT_CustomProperties) */
  readonly customProperties?: readonly CustomPropertyOptions[];
  /** OLE objects (CT_OleObjects) */
  readonly oleObjects?: readonly OleObjectOptions[];
  /** Web publish items (CT_WebPublishItems) */
  readonly webPublishItems?: readonly WebPublishItemOptions[];
}

/** Sheet calc properties (CT_SheetCalcPr) */
export interface SheetCalcPrOptions {
  /** Full calc on load (CT_SheetCalcPr @fullCalcOnLoad) */
  readonly fullCalcOnLoad?: boolean;
}

/** Form control object (CT_Control) */
export interface ControlOptions {
  /** Shape ID (CT_Control @shapeId) */
  readonly shapeId: number;
  /** Control r:id (CT_ControlPr @r:id) */
  readonly rId: string;
  /** Control name (CT_ControlPr @name) */
  readonly name?: string;
  /** Locked (CT_ControlPr @locked) */
  readonly locked?: boolean;
  /** UI-locked (CT_ControlPr @uiObject) */
  readonly uiObject?: boolean;
}

/** Custom property (CT_CustomProperty) */
export interface CustomPropertyOptions {
  /** Property name */
  readonly name: string;
  /** Relationship ID to binary data */
  readonly rId: string;
}

/** OLE object (CT_OleObject) */
export interface OleObjectOptions {
  /** Program ID (CT_OleObject @progId) */
  readonly progId?: string;
  /** Display aspect (CT_OleObject @dvAspect) */
  readonly dvAspect?: "DVASPECT_CONTENT" | "DVASPECT_ICON";
  /** Linked source (CT_OleObject @link) */
  readonly link?: string;
  /** OLE update mode (CT_OleObject @oleUpdate) */
  readonly oleUpdate?: "OLEUPDATE_ALWAYS" | "OLEUPDATE_ONCALL";
  /** Auto load (CT_OleObject @autoLoad) */
  readonly autoLoad?: boolean;
  /** Shape ID (CT_OleObject @shapeId) */
  readonly shapeId: number;
  /** Relationship ID (CT_OleObject @r:id) */
  readonly rId?: string;
  /** Object properties (CT_ObjectPr) */
  readonly objectPr?: OleObjectPrOptions;
}

/** OLE object properties (CT_ObjectPr) */
export interface OleObjectPrOptions {
  /** Locked */
  readonly locked?: boolean;
  /** Default size */
  readonly defaultSize?: boolean;
  /** Print */
  readonly print?: boolean;
  /** Disabled */
  readonly disabled?: boolean;
  /** UI object */
  readonly uiObject?: boolean;
  /** Auto fill */
  readonly autoFill?: boolean;
  /** Auto line */
  readonly autoLine?: boolean;
  /** Auto picture */
  readonly autoPict?: boolean;
  /** Macro */
  readonly macro?: string;
  /** Alt text */
  readonly altText?: string;
  /** DDE */
  readonly dde?: boolean;
  /** Relationship ID */
  readonly rId?: string;
}

/** Web publish item (CT_WebPublishItem) */
export interface WebPublishItemOptions {
  /** Item ID */
  readonly id: number;
  /** HTML div ID */
  readonly divId: string;
  /** Source type */
  readonly sourceType:
    | "sheet"
    | "printArea"
    | "autoFilter"
    | "range"
    | "chart"
    | "pivotTable"
    | "query"
    | "label";
  /** Source cell reference */
  readonly sourceRef?: string;
  /** Source object name */
  readonly sourceObject?: string;
  /** Destination file path */
  readonly destinationFile: string;
  /** Title */
  readonly title?: string;
  /** Auto republish */
  readonly autoRepublish?: boolean;
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
  private readonly rowBreaks: readonly PageBreakOptions[];
  private readonly colBreaks: readonly PageBreakOptions[];
  private readonly customSheetViews: readonly CustomSheetViewOptions[];
  private readonly cellWatches: readonly CellWatchOptions[];
  private readonly dataConsolidate?: DataConsolidateOptions;
  private readonly oleSize?: string;
  private readonly drawingHF?: DrawingHfOptions;
  private readonly legacyDrawingHF?: string;
  private readonly selection?: SelectionOptions;
  private readonly sheetCalcPr?: SheetCalcPrOptions;
  private readonly ext?: string;
  private readonly controls: readonly ControlOptions[];
  private readonly customProperties: readonly CustomPropertyOptions[];
  private readonly oleObjects: readonly OleObjectOptions[];
  private readonly webPublishItems: readonly WebPublishItemOptions[];

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
    this.rowBreaks = options.rowBreaks ?? [];
    this.colBreaks = options.colBreaks ?? [];
    this.customSheetViews = options.customSheetViews ?? [];
    this.cellWatches = options.cellWatches ?? [];
    this.dataConsolidate = options.dataConsolidate;
    this.oleSize = options.oleSize;
    this.drawingHF = options.drawingHF;
    this.legacyDrawingHF = options.legacyDrawingHF;
    this.selection = options.selection;
    this.sheetCalcPr = options.sheetCalcPr;
    this.ext = options.ext;
    this.controls = options.controls ?? [];
    this.customProperties = options.customProperties ?? [];
    this.oleObjects = options.oleObjects ?? [];
    this.webPublishItems = options.webPublishItems ?? [];
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
      // pageSetUpPr (inside sheetPr when fitToPage or autoPageBreaks needed)
      if (this.pageSetup?.fitToWidth || this.pageSetup?.fitToHeight) {
        prParts.push('<pageSetUpPr fitToPage="1"/>');
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
        this.selection ? this.buildSelectionXml(this.selection) : "",
        "</sheetView></sheetViews>",
      );
    } else {
      const svAttrs = this.buildSheetViewAttrs();
      const selStr = this.selection
        ? `>${this.buildSelectionXml(this.selection)}</sheetView>`
        : "/>";
      if (selStr.startsWith(">")) {
        p.push(`<sheetViews><sheetView${svAttrs}${selStr}</sheetViews>`);
      } else {
        p.push(`<sheetViews><sheetView${svAttrs}/></sheetViews>`);
      }
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

    // Sheet calc properties (after sheetData per XSD sequence)
    if (this.sheetCalcPr) {
      const scAttrs: string[] = [];
      if (this.sheetCalcPr.fullCalcOnLoad) scAttrs.push('fullCalcOnLoad="1"');
      p.push(`<sheetCalcPr${scAttrs.length ? " " + scAttrs.join(" ") : ""}/>`);
    }

    // Row breaks (after sheetCalcPr per XSD sequence)
    if (this.rowBreaks.length > 0) {
      const brkParts = this.rowBreaks.map((b) => {
        const bAttrs: Record<string, string | number | boolean | undefined> = { id: b.id };
        if (b.min !== undefined) bAttrs.min = b.min;
        if (b.max !== undefined) bAttrs.max = b.max;
        if (b.manual) bAttrs.man = 1;
        if (b.pivot) bAttrs.pt = 1;
        return `<brk${attrs(bAttrs)}/>`;
      });
      p.push(
        `<rowBreaks count="${this.rowBreaks.length}" manualBreakCount="${this.rowBreaks.filter((b) => b.manual).length}">${brkParts.join("")}</rowBreaks>`,
      );
    }

    // Column breaks
    if (this.colBreaks.length > 0) {
      const brkParts = this.colBreaks.map((b) => {
        const bAttrs: Record<string, string | number | boolean | undefined> = { id: b.id };
        if (b.min !== undefined) bAttrs.min = b.min;
        if (b.max !== undefined) bAttrs.max = b.max;
        if (b.manual) bAttrs.man = 1;
        if (b.pivot) bAttrs.pt = 1;
        return `<brk${attrs(bAttrs)}/>`;
      });
      p.push(
        `<colBreaks count="${this.colBreaks.length}" manualBreakCount="${this.colBreaks.filter((b) => b.manual).length}">${brkParts.join("")}</colBreaks>`,
      );
    }

    // Custom properties (CT_CustomProperties, after colBreaks per XSD sequence)
    if (this.customProperties.length > 0) {
      const cpParts: string[] = ["<customProperties>"];
      for (const cp of this.customProperties) {
        cpParts.push(`<customPr name="${escapeXml(cp.name)}" r:id="${escapeXml(cp.rId)}"/>`);
      }
      cpParts.push("</customProperties>");
      p.push(cpParts.join(""));
    }

    // OLE size
    if (this.oleSize) {
      p.push(`<oleSize ref="${escapeXml(this.oleSize)}"/>`);
    }

    // Custom sheet views (after oleSize per XSD sequence)
    if (this.customSheetViews.length > 0) {
      p.push("<customSheetViews>");
      for (const csv of this.customSheetViews) {
        const csvAttrs: Record<string, string | number | boolean | undefined> = { guid: csv.guid };
        if (csv.scale !== undefined) csvAttrs.scale = csv.scale;
        if (csv.showPageBreaks) csvAttrs.showPageBreaks = 1;
        if (csv.showFormulas) csvAttrs.showFormulas = 1;
        if (csv.showGridLines === false) csvAttrs.showGridLines = 0;
        if (csv.showRowColHeaders === false) csvAttrs.showRowCol = 0;
        if (csv.outlineSymbols === false) csvAttrs.outlineSymbols = 0;
        if (csv.zeroValues === false) csvAttrs.zeroValues = 0;
        if (csv.fitToPage) csvAttrs.fitToPage = 1;
        if (csv.printArea) csvAttrs.printArea = 1;
        if (csv.filter) csvAttrs.filter = 1;
        if (csv.showAutoFilter) csvAttrs.showAutoFilter = 1;
        if (csv.hiddenRows) csvAttrs.hiddenRows = 1;
        if (csv.hiddenColumns) csvAttrs.hiddenColumns = 1;
        if (csv.state && csv.state !== "visible") csvAttrs.state = csv.state;
        if (csv.filterUnique) csvAttrs.filterUnique = 1;
        if (csv.view && csv.view !== "normal") csvAttrs.view = csv.view;
        p.push(`<customSheetView${attrs(csvAttrs)}/>`);
      }
      p.push("</customSheetViews>");
    }

    // Cell watches
    if (this.cellWatches.length > 0) {
      p.push("<cellWatches>");
      for (const cw of this.cellWatches) {
        p.push(`<cellWatch r="${escapeXml(cw.r)}"/>`);
      }
      p.push("</cellWatches>");
    }

    // Data consolidation
    if (this.dataConsolidate) {
      const dc = this.dataConsolidate;
      const dcAttrs: Record<string, string | number | boolean | undefined> = {};
      if (dc.function && dc.function !== "sum") dcAttrs.function = dc.function;
      if (dc.topLabels) dcAttrs.topLabels = 1;
      if (dc.leftLabels) dcAttrs.leftLabels = 1;
      if (dc.startLabels) dcAttrs.startLabels = 1;
      if (dc.link) dcAttrs.link = 1;
      const refsInner = dc.refs?.map((r) => `<dataRef ref="${escapeXml(r)}"/>`).join("") ?? "";
      const refsXml = refsInner ? `<dataRefs>${refsInner}</dataRefs>` : "";
      if (refsXml || Object.keys(dcAttrs).length > 0) {
        p.push(`<dataConsolidate${attrs(dcAttrs)}>${refsXml}</dataConsolidate>`);
      }
    }

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
            if (sc.sortBy) scAttrs.sortBy = sc.sortBy;
            if (sc.customList) scAttrs.customList = sc.customList;
            if (sc.iconId !== undefined) scAttrs.iconId = sc.iconId;
            sortParts.push(selfCloseElement("sortCondition", attrs(scAttrs)));
          }
          const ssAttrs: Record<string, string | number | boolean | undefined> = { ref: af.ref };
          if (af.sortState?.columnSort) ssAttrs.columnSort = 1;
          if (af.sortState?.caseSensitive) ssAttrs.caseSensitive = 1;
          if (af.sortState?.sortMethod) ssAttrs.sortMethod = af.sortState.sortMethod;
          inner.push(`<sortState${attrs(ssAttrs)}>${sortParts.join("")}</sortState>`);
        }
        // Color filters
        for (const cf of af.colorFilters ?? []) {
          const cfAttrs: Record<string, string | number | boolean | undefined> = {};
          if (cf.dxfId !== undefined) cfAttrs.dxfId = cf.dxfId;
          if (cf.cellColor === false) cfAttrs.cellColor = 0;
          inner.push(
            `<filterColumn colId="${cf.colId}"><colorFilter${attrs(cfAttrs)}/></filterColumn>`,
          );
        }
        // Icon filters
        for (const if_ of af.iconFilters ?? []) {
          const ifAttrs: Record<string, string | number | boolean | undefined> = {
            iconSet: if_.iconSet,
          };
          if (if_.iconId !== undefined) ifAttrs.iconId = if_.iconId;
          inner.push(
            `<filterColumn colId="${if_.colId}"><iconFilter${attrs(ifAttrs)}/></filterColumn>`,
          );
        }
        // Dynamic filters
        for (const df of af.dynamicFilters ?? []) {
          const dfAttrs: Record<string, string | number | boolean | undefined> = { type: df.type };
          if (df.val !== undefined) dfAttrs.val = df.val;
          if (df.maxVal !== undefined) dfAttrs.maxVal = df.maxVal;
          inner.push(
            `<filterColumn colId="${df.colId}"><dynamicFilter${attrs(dfAttrs)}/></filterColumn>`,
          );
        }
        // Date group filters
        for (const dg of af.dateGroupItems ?? []) {
          const dgAttrs: Record<string, string | number | boolean | undefined> = {
            dateTimeGrouping: dg.dateTimeGrouping,
          };
          if (dg.year !== undefined) dgAttrs.year = dg.year;
          if (dg.month !== undefined) dgAttrs.month = dg.month;
          if (dg.day !== undefined) dgAttrs.day = dg.day;
          if (dg.hour !== undefined) dgAttrs.hour = dg.hour;
          if (dg.minute !== undefined) dgAttrs.minute = dg.minute;
          if (dg.second !== undefined) dgAttrs.second = dg.second;
          inner.push(
            `<filterColumn colId="${dg.colId}"><dateGroupItem${attrs(dgAttrs)}/></filterColumn>`,
          );
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

    // Drawing in header/footer (after headerFooter per XSD sequence)
    if (this.drawingHF) {
      const dhf = this.drawingHF;
      const dhfAttrs: Record<string, string | number | boolean | undefined> = { "r:id": dhf.rId };
      if (dhf.lho !== undefined) dhfAttrs.lho = dhf.lho;
      if (dhf.lhe !== undefined) dhfAttrs.lhe = dhf.lhe;
      if (dhf.lhf !== undefined) dhfAttrs.lhf = dhf.lhf;
      if (dhf.cho !== undefined) dhfAttrs.cho = dhf.cho;
      if (dhf.che !== undefined) dhfAttrs.che = dhf.che;
      if (dhf.chf !== undefined) dhfAttrs.chf = dhf.chf;
      if (dhf.rho !== undefined) dhfAttrs.rho = dhf.rho;
      if (dhf.rhe !== undefined) dhfAttrs.rhe = dhf.rhe;
      if (dhf.rhf !== undefined) dhfAttrs.rhf = dhf.rhf;
      if (dhf.lfo !== undefined) dhfAttrs.lfo = dhf.lfo;
      if (dhf.lfe !== undefined) dhfAttrs.lfe = dhf.lfe;
      if (dhf.lff !== undefined) dhfAttrs.lff = dhf.lff;
      if (dhf.cfo !== undefined) dhfAttrs.cfo = dhf.cfo;
      if (dhf.cfe !== undefined) dhfAttrs.cfe = dhf.cfe;
      if (dhf.cff !== undefined) dhfAttrs.cff = dhf.cff;
      if (dhf.rfo !== undefined) dhfAttrs.rfo = dhf.rfo;
      if (dhf.rfe !== undefined) dhfAttrs.rfe = dhf.rfe;
      if (dhf.rff !== undefined) dhfAttrs.rff = dhf.rff;
      p.push(selfCloseElement("drawingHF", attrs(dhfAttrs)));
    }

    // Legacy drawing in header/footer
    if (this.legacyDrawingHF) {
      p.push(`<legacyDrawingHF r:id="${escapeXml(this.legacyDrawingHF)}"/>`);
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

    // OLE objects (CT_OleObjects, after picture per XSD sequence)
    if (this.oleObjects.length > 0) {
      const oleParts: string[] = ["<oleObjects>"];
      for (const ole of this.oleObjects) {
        const oleAttrs: string[] = [`shapeId="${ole.shapeId}"`];
        if (ole.progId) oleAttrs.push(`progId="${escapeXml(ole.progId)}"`);
        if (ole.dvAspect && ole.dvAspect !== "DVASPECT_CONTENT")
          oleAttrs.push(`dvAspect="${ole.dvAspect}"`);
        if (ole.link) oleAttrs.push(`link="${escapeXml(ole.link)}"`);
        if (ole.oleUpdate) oleAttrs.push(`oleUpdate="${ole.oleUpdate}"`);
        if (ole.autoLoad) oleAttrs.push('autoLoad="1"');
        if (ole.rId) oleAttrs.push(`r:id="${escapeXml(ole.rId)}"`);
        // objectPr (CT_ObjectPr, optional child)
        if (ole.objectPr) {
          const opr = ole.objectPr;
          const oprAttrs: string[] = [];
          if (opr.locked === false) oprAttrs.push('locked="0"');
          if (opr.defaultSize === false) oprAttrs.push('defaultSize="0"');
          if (opr.print === false) oprAttrs.push('print="0"');
          if (opr.disabled) oprAttrs.push('disabled="1"');
          if (opr.uiObject) oprAttrs.push('uiObject="1"');
          if (opr.autoFill === false) oprAttrs.push('autoFill="0"');
          if (opr.autoLine === false) oprAttrs.push('autoLine="0"');
          if (opr.autoPict === false) oprAttrs.push('autoPict="0"');
          if (opr.macro) oprAttrs.push(`macro="${escapeXml(opr.macro)}"`);
          if (opr.altText) oprAttrs.push(`altText="${escapeXml(opr.altText)}"`);
          if (opr.dde) oprAttrs.push('dde="1"');
          if (opr.rId) oprAttrs.push(`r:id="${escapeXml(opr.rId)}"`);
          oleParts.push(
            `<oleObject ${oleAttrs.join(" ")}><objectPr${oprAttrs.length ? " " + oprAttrs.join(" ") : ""}/></oleObject>`,
          );
        } else {
          oleParts.push(`<oleObject ${oleAttrs.join(" ")}/>`);
        }
      }
      oleParts.push("</oleObjects>");
      p.push(oleParts.join(""));
    }

    // Controls (CT_Controls, after oleObjects per XSD sequence)
    if (this.controls.length > 0) {
      const ctrlParts: string[] = ["<controls>"];
      for (const c of this.controls) {
        const cAttrs: string[] = [`shapeId="${c.shapeId}"`, `r:id="${escapeXml(c.rId)}"`];
        if (c.name) cAttrs.push(`name="${escapeXml(c.name)}"`);
        // controlPr (optional)
        const prAttrs: string[] = [];
        if (c.locked === false) prAttrs.push('locked="0"');
        if (c.uiObject) prAttrs.push('uiObject="1"');
        if (prAttrs.length > 0) {
          ctrlParts.push(
            `<control ${cAttrs.join(" ")}><controlPr${prAttrs.length ? " " + prAttrs.join(" ") : ""}/></control>`,
          );
        } else {
          ctrlParts.push(`<control ${cAttrs.join(" ")}/>`);
        }
      }
      ctrlParts.push("</controls>");
      p.push(ctrlParts.join(""));
    }

    // Web publish items (CT_WebPublishItems, after controls per XSD sequence)
    if (this.webPublishItems.length > 0) {
      const wpParts: string[] = [`<webPublishItems count="${this.webPublishItems.length}">`];
      for (const wpi of this.webPublishItems) {
        const wpiAttrs: string[] = [
          `id="${wpi.id}"`,
          `divId="${escapeXml(wpi.divId)}"`,
          `sourceType="${wpi.sourceType}"`,
          `destinationFile="${escapeXml(wpi.destinationFile)}"`,
        ];
        if (wpi.sourceRef) wpiAttrs.push(`sourceRef="${escapeXml(wpi.sourceRef)}"`);
        if (wpi.sourceObject) wpiAttrs.push(`sourceObject="${escapeXml(wpi.sourceObject)}"`);
        if (wpi.title) wpiAttrs.push(`title="${escapeXml(wpi.title)}"`);
        if (wpi.autoRepublish) wpiAttrs.push('autoRepublish="1"');
        wpParts.push(`<webPublishItem ${wpiAttrs.join(" ")}/>`);
      }
      wpParts.push("</webPublishItems>");
      p.push(wpParts.join(""));
    }

    // Extension list (extLst, last per XSD sequence)
    if (this.ext) {
      p.push(this.ext);
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

  private buildSelectionXml(sel: SelectionOptions): string {
    const selAttrs: Record<string, string | number | boolean | undefined> = {};
    if (sel.pane) selAttrs.pane = sel.pane;
    if (sel.activeCell) selAttrs.activeCell = sel.activeCell;
    if (sel.activeCellId !== undefined) selAttrs.activeCellId = sel.activeCellId;
    if (sel.sqref) selAttrs.sqref = sel.sqref;
    return `<selection${attrs(selAttrs)}/>`;
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

    // Rich text value (RichTextOptions)
    if (typeof value === "object" && !(value instanceof Date)) {
      if (sharedStrings) {
        cellAttrs.t = "s";
        const idx = sharedStrings.registerRich(value);
        return `<c${attrs(cellAttrs)}><v>${idx}</v></c>`;
      }
      cellAttrs.t = "inlineStr";
      return `<c${attrs(cellAttrs)}><is>${buildRstXml(value)}</is></c>`;
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
