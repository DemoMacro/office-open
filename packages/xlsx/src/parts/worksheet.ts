/**
 * Worksheet XML generation — pure functions for xl/worksheets/sheet{n}.xml.
 *
 * All interfaces and the zero-allocation string concatenation fast path
 * are preserved. The `Worksheet` class has been replaced by `buildWorksheetXml()`.
 *
 * @module
 */

import { derivePasswordHash } from "@office-open/core";
import type { ChartSpaceOptions } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attrs, attrsRaw, escapeXml, selfCloseElement } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { findChild, attr, attrNum, textOf } from "@office-open/xml";

import type { XlsxReadContext } from "../context";
import type { PivotTableOptions } from "./pivot";
import { buildRstXml } from "./shared-strings";
import type { SharedStrings } from "./shared-strings";
import type { Styles, StyleOptions } from "./styles";
import type { TableOptions } from "./table";

// ── Option interfaces ──

export interface ColumnOptions {
  min: number;
  max: number;
  width?: number;
  hidden?: boolean;
  customWidth?: boolean;
  outlineLevel?: number;
  collapsed?: boolean;
  /** Best-fit column width (CT_Col @bestFit) */
  bestFit?: boolean;
  /** Phonetic text for CJK (CT_Col @phonetic) */
  phonetic?: boolean;
}

export interface RowOptions {
  cells?: CellOptions[];
  height?: number;
  hidden?: boolean;
  rowNumber?: number;
  /** Spans for the row, e.g. "1:15" (CT_Row @spans) */
  spans?: string;
  /** Custom format applied (CT_Row @customFormat) */
  customFormat?: boolean;
  /** Thick top border (CT_Row @thickTop) */
  thickTop?: boolean;
  /** Thick bottom border (CT_Row @thickBot) */
  thickBot?: boolean;
  /** Phonetic text (CT_Row @ph) */
  ph?: boolean;
}

/** Rich text run properties (CT_RPrElt). */
export interface RichTextRunPrOptions {
  /** Font name (CT_FontName → rFont) */
  font?: string;
  /** Character set (CT_IntProperty) */
  charset?: number;
  /** Font family (CT_IntProperty) */
  family?: number;
  /** Bold */
  bold?: boolean;
  /** Italic */
  italic?: boolean;
  /** Strikethrough */
  strike?: boolean;
  /** Outline */
  outline?: boolean;
  /** Shadow */
  shadow?: boolean;
  /** Condense */
  condense?: boolean;
  /** Extend */
  extend?: boolean;
  /** Font color (hex RGB, e.g. "FF0000") */
  color?: string;
  /** Font size in points */
  size?: number;
  /** Underline type */
  underline?: "single" | "double" | "singleAccounting" | "doubleAccounting" | "none";
  /** Vertical alignment */
  vertAlign?: "superscript" | "subscript" | "baseline";
  /** Font scheme */
  scheme?: "major" | "minor" | "none";
}

/** A single rich text run (CT_RElt). */
export interface RichTextRunOptions {
  /** Run properties (optional = inherits from parent) */
  properties?: RichTextRunPrOptions;
  /** Run text content */
  text: string;
}

/** Phonetics run for CJK (CT_PhoneticRun → rPh). */
export interface PhoneticRunOptions {
  /** Start byte offset in base text */
  sb: number;
  /** End byte offset in base text */
  eb: number;
  /** Phonetic text */
  text: string;
}

/** Rich text content (CT_Rst). Either plain text or rich runs. */
export interface RichTextOptions {
  /** Plain text (mutually exclusive with runs) */
  text?: string;
  /** Rich text runs (mutually exclusive with text) */
  runs?: RichTextRunOptions[];
  /** Phonetic runs for CJK (CT_PhoneticRun) */
  phonetics?: PhoneticRunOptions[];
}

export interface CellOptions {
  value?: string | number | boolean | Date | RichTextOptions | null;
  reference?: string;
  /** Direct style index (for pre-resolved styles) */
  styleIndex?: number;
  /** Style options (resolved to index at compile time) */
  style?: StyleOptions;
  /** Formula options. When set, value becomes the cached result. */
  formula?: FormulaOptions;
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
  formula: string;
  /** Formula type (default: "normal") */
  type?: FormulaType;
  /** Reference range for array/shared formulas, e.g. "C1:C10" */
  reference?: string;
  /** Shared formula group index (required for shared formulas) */
  sharedIndex?: number;
  /** Always calculate array (CT_CellFormula @aca) */
  aca?: boolean;
  /** 2-D data table (CT_CellFormula @dt2D) */
  dt2D?: boolean;
  /** Data table row (CT_CellFormula @dtr) */
  dtr?: boolean;
  /** Delete input cell 1 (CT_CellFormula @del1) */
  del1?: boolean;
  /** Delete input cell 2 (CT_CellFormula @del2) */
  del2?: boolean;
  /** Input cell 1 reference (CT_CellFormula @r1) */
  r1?: string;
  /** Input cell 2 reference (CT_CellFormula @r2) */
  r2?: string;
  /** Calculate cell (CT_CellFormula @ca) */
  ca?: boolean;
  /** Array formula context (CT_CellFormula @bx) */
  bx?: boolean;
}

/** Input cell for a what-if scenario (maps to CT_InputCells). */
export interface ScenarioCellOptions {
  /** Cell reference, e.g. "B2" */
  r: string;
  /** Cell value for this scenario */
  val: string | number;
  /** Whether the value is deleted */
  deleted?: boolean;
  /** Whether undone (CT_InputCells @undone) */
  undone?: boolean;
}

/** A single what-if scenario (maps to CT_Scenario). */
export interface ScenarioDefinition {
  /** Scenario name */
  name: string;
  /** Input cells with their values for this scenario */
  inputCells: ScenarioCellOptions[];
  /** Sort/order count */
  count?: number;
  /** Creator user name */
  user?: string;
  /** Comment */
  comment?: string;
  /** Whether the scenario is hidden */
  hidden?: boolean;
  /** Whether the scenario is locked */
  locked?: boolean;
}

/** Scenarios for what-if analysis (maps to CT_Scenarios). */
export interface ScenarioOptions {
  /** Named scenarios */
  scenarios: ScenarioDefinition[];
  /** Current scenario index (0-based) */
  current?: number;
  /** Show scenario index (0-based) */
  show?: number;
}

export interface MergeCellOptions {
  from: { row: number; col: number };
  to: { row: number; col: number };
}

export interface SheetProtectionOptions {
  /** Plain-text password — legacy Excel hash is computed automatically */
  password?: string;
  /** Modern encryption: algorithm name (e.g. "SHA-512") */
  algorithmName?: string;
  /** Modern encryption: base64-encoded hash value */
  hashValue?: string;
  /** Modern encryption: base64-encoded salt value */
  saltValue?: string;
  /** Modern encryption: spin count for hash iteration */
  spinCount?: number;
  /** Set true to enable sheet protection (required for protection flags to take effect) */
  sheet?: boolean;
  objects?: boolean;
  scenarios?: boolean;
  formatCells?: boolean;
  formatColumns?: boolean;
  formatRows?: boolean;
  insertColumns?: boolean;
  insertRows?: boolean;
  insertHyperlinks?: boolean;
  deleteColumns?: boolean;
  deleteRows?: boolean;
  selectLockedCells?: boolean;
  sort?: boolean;
  autoFilter?: boolean;
  pivotTables?: boolean;
  selectUnlockedCells?: boolean;
}

/** A named protected range within a sheet (CT_ProtectedRange) */
export interface ProtectedRangeOptions {
  /** Range reference (required), e.g. "A1:C10" */
  sqref: string;
  /** Range name (required) */
  name: string;
  /** Plain-text password — legacy hash computed automatically */
  password?: string;
  /** Modern encryption: algorithm name */
  algorithmName?: string;
  /** Modern encryption: base64-encoded hash value */
  hashValue?: string;
  /** Modern encryption: base64-encoded salt value */
  saltValue?: string;
  /** Modern encryption: spin count */
  spinCount?: number;
  /** Security descriptor (SID string) */
  securityDescriptor?: string;
}

export interface FreezePaneOptions {
  /** Row split position (1-based, freezes rows above) */
  row?: number;
  /** Column split position (1-based, freezes columns to the left) */
  col?: number;
}

export interface WorksheetImageOptions {
  data: Uint8Array;
  type: "png" | "jpeg" | "jpg";
  col: number;
  row: number;
}

export interface WorksheetChartOptions extends ChartSpaceOptions {
  /** 1-based column position for the chart */
  col: number;
  /** 1-based row position for the chart */
  row: number;
}

export interface SheetViewOptions {
  showGridLines?: boolean;
  showRowColHeaders?: boolean;
  showZeros?: boolean;
  zoomScale?: number;
  tabSelected?: boolean;
  rightToLeft?: boolean;
  /** Window protection (CT_SheetView @windowProtection) */
  windowProtection?: boolean;
  /** Show formulas instead of values (CT_SheetView @showFormulas) */
  showFormulas?: boolean;
  /** Show ruler (CT_SheetView @showRuler) */
  showRuler?: boolean;
  /** Show outline symbols (CT_SheetView @showOutlineSymbols) */
  showOutlineSymbols?: boolean;
  /** Default grid color (CT_SheetView @defaultGridColor) */
  defaultGridColor?: boolean;
  /** Show white space (CT_SheetView @showWhiteSpace) */
  showWhiteSpace?: boolean;
  /** View type (CT_SheetView @view) */
  view?: "normal" | "pageBreakPreview" | "pageLayout";
  /** Tab color ID (CT_SheetView @colorId) */
  colorId?: number;
  /** Zoom scale for normal view (CT_SheetView @zoomScaleNormal) */
  zoomScaleNormal?: number;
  /** Zoom scale for sheet layout view (CT_SheetView @zoomScaleSheetLayoutView) */
  zoomScaleSheetLayoutView?: number;
  /** Zoom scale for page layout view (CT_SheetView @zoomScalePageLayoutView) */
  zoomScalePageLayoutView?: number;
  /** Pivot selections (CT_PivotSelection) */
  pivotSelections?: PivotSelectionOptions[];
}

/** Pivot selection in sheet view (CT_PivotSelection) */
export interface PivotSelectionOptions {
  /** Pane (default: "topLeft") */
  pane?: "bottomRight" | "topRight" | "bottomLeft" | "topLeft";
  /** Show header (default: false) */
  showHeader?: boolean;
  /** Label selected (default: false) */
  label?: boolean;
  /** Data selected (default: false) */
  data?: boolean;
  /** Extendable (default: false) */
  extendable?: boolean;
  /** Selection count */
  count?: number;
  /** Axis */
  axis?: "axisRow" | "axisCol" | "axisPage" | "axisValues";
  /** Dimension */
  dimension?: number;
  /** Start index */
  start?: number;
  /** Min index */
  min?: number;
  /** Max index */
  max?: number;
  /** Active row */
  activeRow?: number;
  /** Active column */
  activeCol?: number;
  /** Previous row */
  previousRow?: number;
  /** Previous column */
  previousCol?: number;
  /** Clicked row */
  click?: number;
  /** Relationship ID (maps to r:id in XML) */
  rId?: string;
}

export type HyperlinkTarget =
  | { type: "external"; url: string }
  | { type: "internal"; location: string };

export interface HyperlinkOptions {
  /** Cell reference, e.g. "A1" */
  cell: string;
  /** Hyperlink target */
  target: HyperlinkTarget;
  /** Tooltip text */
  tooltip?: string;
  /** Display text */
  display?: string;
}

export interface HeaderFooterOptions {
  oddHeader?: string;
  oddFooter?: string;
  evenHeader?: string;
  evenFooter?: string;
  firstHeader?: string;
  firstFooter?: string;
  differentOddEven?: boolean;
  differentFirst?: boolean;
  /** Scale header/footer with document (CT_HeaderFooter @scaleWithDoc) */
  scaleWithDoc?: boolean;
  /** Align with page margins (CT_HeaderFooter @alignWithMargins) */
  alignWithMargins?: boolean;
}

export type PageOrientation = "default" | "portrait" | "landscape";

export interface PageSetupOptions {
  paperSize?: number;
  orientation?: PageOrientation;
  scale?: number;
  fitToWidth?: number;
  fitToHeight?: number;
  pageOrder?: "downThenOver" | "overThenDown";
  useFirstPageNumber?: boolean;
  firstPageNumber?: number;
  /** Paper height (CT_PageSetup @paperHeight) */
  paperHeight?: number;
  /** Paper width (CT_PageSetup @paperWidth) */
  paperWidth?: number;
  /** Use printer defaults (CT_PageSetup @usePrinterDefaults) */
  usePrinterDefaults?: boolean;
  /** Black and white printing (CT_PageSetup @blackAndWhite) */
  blackAndWhite?: boolean;
  /** Draft quality printing (CT_PageSetup @draft) */
  draft?: boolean;
  /** Print cell comments mode (CT_PageSetup @cellComments) */
  cellComments?: "none" | "asDisplayed" | "atEnd";
  /** Print error display mode (CT_PageSetup @errors) */
  errors?: "displayed" | "blank" | "dash" | "NA";
  /** Auto page breaks (CT_PageSetUpPr @autoPageBreaks) */
  autoPageBreaks?: boolean;
  /** Fit to page (CT_PageSetUpPr @fitToPage) */
  fitToPage?: boolean;
}

export interface TabColorOptions {
  /** RGB color string, e.g. "FF0000" */
  rgb?: string;
  /** Theme color index (0-based) */
  theme?: number;
  /** Tint value (-1.0 to 1.0) */
  tint?: number;
  /** Indexed color (CT_Color @indexed) */
  indexed?: number;
}

/** Object anchor (CT_ObjectAnchor). */
export interface ObjectAnchorOptions {
  /** Move with cells (default: false) */
  moveWithCells?: boolean;
  /** Size with cells (default: false) */
  sizeWithCells?: boolean;
}

/** Comment property (CT_CommentPr). */
export interface CommentPrOptions {
  /** Locked */
  locked?: boolean;
  /** Default size */
  defaultSize?: boolean;
  /** Print */
  print?: boolean;
  /** Disabled */
  disabled?: boolean;
  /** Auto fill */
  autoFill?: boolean;
  /** Auto line */
  autoLine?: boolean;
  /** Alt text */
  altText?: string;
  /** Text horizontal alignment */
  textHAlign?: "left" | "center" | "right" | "justify" | "distributed";
  /** Text vertical alignment */
  textVAlign?: "top" | "center" | "bottom" | "justify" | "distributed";
  /** Lock text */
  lockText?: boolean;
  /** Justify last line */
  justLastX?: boolean;
  /** Auto scale */
  autoScale?: boolean;
  /** Object anchor position */
  anchor?: ObjectAnchorOptions;
}

export interface CommentOptions {
  /** Cell reference, e.g. "A1" */
  cell: string;
  /** Author name */
  author: string;
  /** Comment text (plain string or rich text) */
  text: string | RichTextOptions;
  /** Comment properties (CT_CommentPr) */
  commentPr?: CommentPrOptions;
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
  sqref: string;
  type?: DataValidationType;
  operator?: DataValidationOperator;
  formula1?: string;
  formula2?: string;
  allowBlank?: boolean;
  showErrorMessage?: boolean;
  errorTitle?: string;
  error?: string;
  showInputMessage?: boolean;
  promptTitle?: string;
  prompt?: string;
  /** Error style (CT_DataValidation @errorStyle) */
  errorStyle?: "stop" | "warning" | "information";
  /** IME mode (CT_DataValidation @imeMode) */
  imeMode?:
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
  showDropDown?: boolean;
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
  type: CfvoType;
  val?: string | number;
  /** Greater than or equal (default: true) */
  gte?: boolean;
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
  cfvo: CfvoOptions[];
  /** Colors for each value (same count as cfvo) — RGB hex without alpha, e.g. "FF0000" */
  colors: string[];
}

/** Data bar rule configuration */
export interface DataBarOptions {
  /** Minimum and maximum value objects (exactly 2) */
  cfvo: [CfvoOptions, CfvoOptions];
  /** Bar color — RGB hex without alpha, e.g. "638EC6" */
  color: string;
  /** Minimum bar length as percentage (default: 10) */
  minLength?: number;
  /** Maximum bar length as percentage (default: 90) */
  maxLength?: number;
  /** Whether to show cell values (default: true) */
  showValue?: boolean;
}

/** Icon set rule configuration */
export interface IconSetOptions {
  /** Conditional format values (minimum 2) */
  cfvo: CfvoOptions[];
  /** Icon set type (default: "3TrafficLights1") */
  iconSet?: IconSetType;
  /** Whether to show cell values (default: true) */
  showValue?: boolean;
  /** Whether values are percentages (default: true) */
  percent?: boolean;
  /** Whether to reverse icon order (default: false) */
  reverse?: boolean;
}

export interface ConditionalFormatRule {
  type: ConditionalFormatType;
  operator?: ConditionalFormatOperator;
  /** Formula(s) — up to 3 */
  formulas?: string[];
  priority?: number;
  /** Reference to a dxf (differential format) in the styles table */
  dxfId?: number;
  /** Color scale configuration (when type is "colorScale") */
  colorScale?: ColorScaleOptions;
  /** Data bar configuration (when type is "dataBar") */
  dataBar?: DataBarOptions;
  /** Icon set configuration (when type is "iconSet") */
  iconSet?: IconSetOptions;
  /** Stop if true — skip remaining rules (CT_CfRule @stopIfTrue) */
  stopIfTrue?: boolean;
  /** Time period for date-based highlighting (CT_CfRule @timePeriod) */
  timePeriod?:
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
  rank?: number;
  /** Equal average flag (CT_CfRule @equalAverage) */
  equalAverage?: boolean;
}

export interface ConditionalFormatOptions {
  /** Cell range, e.g. "A1:A10" */
  sqref: string;
  rules: ConditionalFormatRule[];
}

export interface Top10FilterOptions {
  colId: number;
  top?: boolean;
  percent?: boolean;
  val: number;
  /** Filter value (CT_Top10 @filterVal) */
  filterVal?: number;
  /** Hide auto-filter button (CT_FilterColumn @hiddenButton) */
  hiddenButton?: boolean;
  /** Show filter button (CT_FilterColumn @showButton) */
  showButton?: boolean;
}

export interface CustomFilterOptions {
  colId: number;
  operator?:
    | "equal"
    | "notEqual"
    | "greaterThan"
    | "greaterThanOrEqual"
    | "lessThan"
    | "lessThanOrEqual";
  val?: string;
  and?: boolean;
  val2?: string;
  /** Hide auto-filter button (CT_FilterColumn @hiddenButton) */
  hiddenButton?: boolean;
  /** Show filter button (CT_FilterColumn @showButton) */
  showButton?: boolean;
}

export interface SortCondition {
  /** Cell reference for the sort column, e.g. "B1" */
  ref: string;
  descending?: boolean;
  /** Sort by (CT_SortCondition @sortBy) */
  sortBy?: "value" | "cellColor" | "fontColor" | "icon";
  /** Custom sort list (CT_SortCondition @customList) */
  customList?: string;
  /** Icon set index (CT_SortCondition @iconId) */
  iconId?: number;
}

export interface AutoFilterOptions {
  /** Range, e.g. "A1:D10" */
  ref: string;
  top10?: Top10FilterOptions[];
  customFilters?: CustomFilterOptions[];
  sort?: SortCondition[];
  /** Sort state options */
  sortState?: SortStateOptions;
  /** Color filters (CT_ColorFilter) */
  colorFilters?: ColorFilterOptions[];
  /** Icon filters (CT_IconFilter) */
  iconFilters?: IconFilterOptions[];
  /** Dynamic filters (CT_DynamicFilter) */
  dynamicFilters?: DynamicFilterOptions[];
  /** Date group items in filters (CT_DateGroupItem) */
  dateGroupItems?: DateGroupFilterOptions[];
  /** Simple filters with values (CT_Filters) */
  filters?: FilterItemsOptions[];
}

/** Color filter (CT_ColorFilter) */
export interface ColorFilterOptions {
  /** Column ID */
  colId: number;
  /** Cell color RGB (dxfId used if not set) */
  dxfId?: number;
  /** Filter by cell color (CT_ColorFilter @cellColor) */
  cellColor?: boolean;
}

/** Icon filter (CT_IconFilter) */
export interface IconFilterOptions {
  /** Column ID */
  colId: number;
  /** Icon set index (CT_IconFilter @iconSet) */
  iconSet: number;
  /** Icon ID within set (CT_IconFilter @iconId) */
  iconId?: number;
}

/** Filter items (CT_Filters) */
export interface FilterItemsOptions {
  /** Column ID */
  colId: number;
  /** Blank filter (CT_Filters @blank) */
  blank?: boolean;
  /** Calendar type (CT_Filters @calendarType) */
  calendarType?: string;
  /** Filter values */
  values?: string[];
}

/** Dynamic filter (CT_DynamicFilter) */
export interface DynamicFilterOptions {
  /** Column ID */
  colId: number;
  /** Dynamic filter type (CT_DynamicFilter @type) */
  type:
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
  val?: number;
  /** Max value as date ISO string (CT_DynamicFilter @maxVal) */
  maxVal?: number;
  /** Value ISO date string (CT_DynamicFilter @valIso) */
  valIso?: string;
  /** Max value ISO date string (CT_DynamicFilter @maxValIso) */
  maxValIso?: string;
}

/** Date group filter item (CT_DateGroupItem) */
export interface DateGroupFilterOptions {
  /** Column ID */
  colId: number;
  /** Date grouping level (CT_DateGroupItem @dateTimeGrouping) */
  dateTimeGrouping: "year" | "month" | "day" | "hour" | "minute" | "second";
  /** Year (CT_DateGroupItem @year) */
  year?: number;
  /** Month (1-12, CT_DateGroupItem @month) */
  month?: number;
  /** Day (1-31, CT_DateGroupItem @day) */
  day?: number;
  /** Hour (0-23, CT_DateGroupItem @hour) */
  hour?: number;
  /** Minute (0-59, CT_DateGroupItem @minute) */
  minute?: number;
  /** Second (0-59, CT_DateGroupItem @second) */
  second?: number;
}

/** Sort state configuration (CT_SortState) */
export interface SortStateOptions {
  /** Column sort mode (CT_SortState @columnSort) */
  columnSort?: boolean;
  /** Case sensitive sorting (CT_SortState @caseSensitive) */
  caseSensitive?: boolean;
  /** Sort method (CT_SortState @sortMethod) */
  sortMethod?: "pinYin" | "stroke";
}

/** Print options (CT_PrintOptions) */
export interface PrintOptions {
  /** Center horizontally on page */
  horizontalCentered?: boolean;
  /** Center vertically on page */
  verticalCentered?: boolean;
  /** Print row/column headings */
  headings?: boolean;
  /** Print grid lines */
  gridLines?: boolean;
  /** Grid lines set flag */
  gridLinesSet?: boolean;
}

/** Sheet format properties (CT_SheetFormatPr) */
export interface SheetFormatPrOptions {
  /** Base column width (CT_SheetFormatPr @baseColWidth) */
  baseColWidth?: number;
  /** Default column width (CT_SheetFormatPr @defaultColWidth) */
  defaultColWidth?: number;
  /** Default row height */
  defaultRowHeight?: number;
  /** Zero height rows hidden (CT_SheetFormatPr @zeroHeight) */
  zeroHeight?: boolean;
  /** Thick top borders (CT_SheetFormatPr @thickTop) */
  thickTop?: boolean;
  /** Thick bottom borders (CT_SheetFormatPr @thickBottom) */
  thickBottom?: boolean;
  /** Outline level row (CT_SheetFormatPr @outlineLevelRow) */
  outlineLevelRow?: number;
  /** Outline level column (CT_SheetFormatPr @outlineLevelCol) */
  outlineLevelCol?: number;
}

/** Sheet properties extended options (CT_SheetPr attributes) */
export interface SheetPrOptions {
  /** Sync horizontal scroll (CT_SheetPr @syncHorizontal) */
  syncHorizontal?: boolean;
  /** Sync vertical scroll (CT_SheetPr @syncVertical) */
  syncVertical?: boolean;
  /** Sync reference (CT_SheetPr @syncRef) */
  syncRef?: string;
  /** Transition evaluation mode (CT_SheetPr @transitionEvaluation) */
  transitionEvaluation?: boolean;
  /** Transition entry mode (CT_SheetPr @transitionEntry) */
  transitionEntry?: boolean;
  /** Published to server (CT_SheetPr @published) */
  published?: boolean;
  /** Filter mode (CT_SheetPr @filterMode) */
  filterMode?: boolean;
  /** Enable format conditions calculation (CT_SheetPr @enableFormatConditionsCalculation) */
  enableFormatConditionsCalculation?: boolean;
  /** Outline apply styles (CT_OutlinePr @applyStyles) */
  outlineApplyStyles?: boolean;
  /** Outline show symbols (CT_OutlinePr @showOutlineSymbols) */
  outlineShowSymbols?: boolean;
  /** Outline summary rows below detail (CT_OutlinePr @summaryBelow) */
  outlineSummaryBelow?: boolean;
  /** Outline summary columns right of detail (CT_OutlinePr @summaryRight) */
  outlineSummaryRight?: boolean;
}

/** An ignored error entry — suppresses specific Excel error checks for a range. */
export interface IgnoredErrorOptions {
  /** Cell range, e.g. "A1:A10" (required) */
  sqref: string;
  evalError?: boolean;
  twoDigitTextYear?: boolean;
  numberStoredAsText?: boolean;
  formula?: boolean;
  formulaRange?: boolean;
  unlockedFormula?: boolean;
  emptyCellReference?: boolean;
  listDataValidation?: boolean;
  calculatedColumn?: boolean;
}

/** Phonetic properties for CJK text (CT_PhoneticPr) */
export interface PhoneticPrOptions {
  /** Font ID from the styles table (required) */
  fontId: number;
  /** Phonetic type (default: "fullwidthKatakana") */
  type?: "fullwidthKatakana" | "halfwidthKatakana" | "Hiragana" | "noConversion";
  /** Alignment (default: "left") */
  alignment?: "left" | "center" | "distributed";
}

/** Background image for a worksheet */
export interface SheetBackgroundImageOptions {
  data: Uint8Array;
  type: "png" | "jpeg" | "jpg";
}

/** Page break entry (CT_Break) */
export interface PageBreakOptions {
  /** Row or column ID (1-based) */
  id: number;
  /** Min value (CT_Break @min) */
  min?: number;
  /** Max value (CT_Break @max) */
  max?: number;
  /** Manual break (CT_Break @man) */
  manual?: boolean;
  /** Pivot break (CT_Break @pt) */
  pivot?: boolean;
}

/** Selection in sheet view (CT_Selection) */
export interface SelectionOptions {
  /** Pane (CT_Selection @pane) */
  pane?: "bottomRight" | "topRight" | "bottomLeft" | "topLeft";
  /** Active cell (CT_Selection @activeCell) */
  activeCell?: string;
  /** Active cell index (CT_Selection @activeCellId) */
  activeCellId?: number;
  /** Selected range (CT_Selection @sqref) */
  sqref?: string;
}

/** Custom sheet view (CT_CustomSheetView) */
export interface CustomSheetViewOptions {
  /** GUID identifier (required, CT_CustomSheetView @guid) */
  guid: string;
  /** Zoom scale (CT_CustomSheetView @scale) */
  scale?: number;
  /** Show page breaks (CT_CustomSheetView @showPageBreaks) */
  showPageBreaks?: boolean;
  /** Show formulas (CT_CustomSheetView @showFormulas) */
  showFormulas?: boolean;
  /** Show grid lines (CT_CustomSheetView @showGridLines) */
  showGridLines?: boolean;
  /** Show row/column headers (CT_CustomSheetView @showRowCol) */
  showRowColHeaders?: boolean;
  /** Show outline symbols (CT_CustomSheetView @outlineSymbols) */
  outlineSymbols?: boolean;
  /** Show zero values (CT_CustomSheetView @zeroValues) */
  zeroValues?: boolean;
  /** Fit to page (CT_CustomSheetView @fitToPage) */
  fitToPage?: boolean;
  /** Print area (CT_CustomSheetView @printArea) */
  printArea?: boolean;
  /** Filter applied (CT_CustomSheetView @filter) */
  filter?: boolean;
  /** Show auto filter (CT_CustomSheetView @showAutoFilter) */
  showAutoFilter?: boolean;
  /** Hidden rows (CT_CustomSheetView @hiddenRows) */
  hiddenRows?: boolean;
  /** Hidden columns (CT_CustomSheetView @hiddenColumns) */
  hiddenColumns?: boolean;
  /** Sheet state (CT_CustomSheetView @state) */
  state?: "visible" | "hidden" | "veryHidden";
  /** Filter unique (CT_CustomSheetView @filterUnique) */
  filterUnique?: boolean;
  /** View type (CT_CustomSheetView @view) */
  view?: "normal" | "pageBreakPreview" | "pageLayout";
}

/** Cell watch entry (CT_CellWatch) */
export interface CellWatchOptions {
  /** Cell reference, e.g. "A1" */
  r: string;
}

/** Data consolidation (CT_DataConsolidate) */
export interface DataConsolidateOptions {
  /** Consolidation function (CT_DataConsolidate @function) */
  function?:
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
  topLabels?: boolean;
  /** Use left column labels (CT_DataConsolidate @leftLabels) */
  leftLabels?: boolean;
  /** Use labels in first row (CT_DataConsolidate @startLabels alias) */
  startLabels?: boolean;
  /** Link to source data (CT_DataConsolidate @link) */
  link?: boolean;
  /** Source data references */
  refs?: string[];
}

/** Drawing in header/footer (CT_DrawingHF) */
export interface DrawingHfOptions {
  /** Relationship ID for the drawing (required) */
  rId: string;
  lho?: number;
  lhe?: number;
  lhf?: number;
  cho?: number;
  che?: number;
  chf?: number;
  rho?: number;
  rhe?: number;
  rhf?: number;
  lfo?: number;
  lfe?: number;
  lff?: number;
  cfo?: number;
  cfe?: number;
  cff?: number;
  rfo?: number;
  rfe?: number;
  rff?: number;
}

export interface WorksheetOptions {
  name?: string;
  rows?: RowOptions[];
  columns?: ColumnOptions[];
  mergeCells?: MergeCellOptions[];
  freezePanes?: FreezePaneOptions;
  protection?: SheetProtectionOptions;
  /** Named protected ranges within this sheet */
  protectedRanges?: ProtectedRangeOptions[];
  /** What-if scenarios */
  scenarios?: ScenarioOptions;
  /** Auto-filter configuration */
  autoFilter?: string | AutoFilterOptions;
  images?: WorksheetImageOptions[];
  charts?: WorksheetChartOptions[];
  dataValidations?: DataValidationOptions[];
  /** Disable data validation prompts (CT_DataValidations @disablePrompts) */
  dataValidationsDisablePrompts?: boolean;
  conditionalFormats?: ConditionalFormatOptions[];
  hyperlinks?: HyperlinkOptions[];
  comments?: CommentOptions[];
  headerFooter?: HeaderFooterOptions;
  pageSetup?: PageSetupOptions;
  tabColor?: TabColorOptions;
  sheetView?: SheetViewOptions;
  pivotTables?: PivotTableOptions[];
  /** Tables (list objects) for this worksheet */
  tables?: TableOptions[];
  /** Ignored errors — suppress specific Excel error checks for cell ranges */
  ignoredErrors?: IgnoredErrorOptions[];
  /** Phonetic properties for CJK text */
  phoneticPr?: PhoneticPrOptions;
  /** Background image for the worksheet */
  backgroundImage?: SheetBackgroundImageOptions;
  /** Print options (CT_PrintOptions) */
  printOptions?: PrintOptions;
  /** Sheet format properties (CT_SheetFormatPr) */
  sheetFormatPr?: SheetFormatPrOptions;
  /** Sheet extended properties (CT_SheetPr attributes) */
  sheetPr?: SheetPrOptions;
  /** Row page breaks (CT_PageBreaks) */
  rowBreaks?: PageBreakOptions[];
  /** Column page breaks (CT_PageBreaks) */
  colBreaks?: PageBreakOptions[];
  /** Custom sheet views (CT_CustomSheetViews) */
  customSheetViews?: CustomSheetViewOptions[];
  /** Cell watches (CT_CellWatches) */
  cellWatches?: CellWatchOptions[];
  /** Data consolidation (CT_DataConsolidate) */
  dataConsolidate?: DataConsolidateOptions;
  /** OLE embedded range (CT_OleSize) */
  oleSize?: string;
  /** Drawing in header/footer (CT_DrawingHF) */
  drawingHF?: DrawingHfOptions;
  /** Legacy drawing for header/footer r:id (CT_LegacyDrawingHF) */
  legacyDrawingHF?: string;
  /** Selection in sheet view (CT_Selection) */
  selection?: SelectionOptions;
  /** Sheet calc properties (CT_SheetCalcPr) */
  sheetCalcPr?: SheetCalcPrOptions;
  /** Extension list (extLst) */
  ext?: string;
  /** Control objects (CT_Controls) */
  controls?: ControlOptions[];
  /** Custom sheet properties (CT_CustomProperties) */
  customProperties?: CustomPropertyOptions[];
  /** OLE objects (CT_OleObjects) */
  oleObjects?: OleObjectOptions[];
  /** Web publish items (CT_WebPublishItems) */
  webPublishItems?: WebPublishItemOptions[];
}

/** Sheet calc properties (CT_SheetCalcPr) */
export interface SheetCalcPrOptions {
  /** Full calc on load (CT_SheetCalcPr @fullCalcOnLoad) */
  fullCalcOnLoad?: boolean;
}

/** Form control object (CT_Control) */
export interface ControlOptions {
  /** Shape ID (CT_Control @shapeId) */
  shapeId: number;
  /** Control r:id (CT_ControlPr @r:id) */
  rId: string;
  /** Control name (CT_ControlPr @name) */
  name?: string;
  /** Locked (CT_ControlPr @locked) */
  locked?: boolean;
  /** UI-locked (CT_ControlPr @uiObject) */
  uiObject?: boolean;
  /** Recalc always (CT_ControlPr @recalcAlways) */
  recalcAlways?: boolean;
  /** Linked cell (CT_ControlPr @linkedCell) */
  linkedCell?: string;
  /** List fill range (CT_ControlPr @listFillRange) */
  listFillRange?: string;
  /** Control formula (CT_ControlPr @cf) */
  cf?: string;
}

/** Custom property (CT_CustomProperty) */
export interface CustomPropertyOptions {
  /** Property name */
  name: string;
  /** Relationship ID to binary data */
  rId: string;
}

/** OLE object (CT_OleObject) */
export interface OleObjectOptions {
  /** Program ID (CT_OleObject @progId) */
  progId?: string;
  /** Display aspect (CT_OleObject @dvAspect) */
  dvAspect?: "DVASPECT_CONTENT" | "DVASPECT_ICON";
  /** Linked source (CT_OleObject @link) */
  link?: string;
  /** OLE update mode (CT_OleObject @oleUpdate) */
  oleUpdate?: "OLEUPDATE_ALWAYS" | "OLEUPDATE_ONCALL";
  /** Auto load (CT_OleObject @autoLoad) */
  autoLoad?: boolean;
  /** Shape ID (CT_OleObject @shapeId) */
  shapeId: number;
  /** Relationship ID (CT_OleObject @r:id) */
  rId?: string;
  /** Object properties (CT_ObjectPr) */
  objectPr?: OleObjectPrOptions;
}

/** OLE object properties (CT_ObjectPr) */
export interface OleObjectPrOptions {
  /** Locked */
  locked?: boolean;
  /** Default size */
  defaultSize?: boolean;
  /** Print */
  print?: boolean;
  /** Disabled */
  disabled?: boolean;
  /** UI object */
  uiObject?: boolean;
  /** Auto fill */
  autoFill?: boolean;
  /** Auto line */
  autoLine?: boolean;
  /** Auto picture */
  autoPict?: boolean;
  /** Macro */
  macro?: string;
  /** Alt text */
  altText?: string;
  /** DDE */
  dde?: boolean;
  /** Relationship ID */
  rId?: string;
}

/** Web publish item (CT_WebPublishItem) */
export interface WebPublishItemOptions {
  /** Item ID */
  id: number;
  /** HTML div ID */
  divId: string;
  /** Source type */
  sourceType:
    | "sheet"
    | "printArea"
    | "autoFilter"
    | "range"
    | "chart"
    | "pivotTable"
    | "query"
    | "label";
  /** Source cell reference */
  sourceRef?: string;
  /** Source object name */
  sourceObject?: string;
  /** Destination file path */
  destinationFile: string;
  /** Title */
  title?: string;
  /** Auto republish */
  autoRepublish?: boolean;
}

// ── Worksheet XML builder context ──

/** Minimal context needed by buildWorksheetXml. */
export interface WorksheetContext {
  sharedStrings?: SharedStrings;
  styles?: Styles;
}

// ── Pure functions ──

// Re-exported for use by the compiler (defined below in this file).
export { stringifyWorksheet as buildWorksheetXml };
// ── Descriptor ──

export const worksheetDesc: CustomDescriptor<WorksheetOptions> = {
  kind: "custom",

  /**
   * NOT intended for direct use by the compiler.
   * The compiler calls `stringifyWorksheet(opts, ctx)` instead, which has
   * access to the SharedStrings and Styles accumulators.
   * This method exists to satisfy the CustomDescriptor interface for the read path.
   */
  stringify(_opts, _ctx) {
    throw new Error(
      "Use stringifyWorksheet(opts, ctx) for the write path. worksheetDesc.stringify() is not supported.",
    );
  },

  parse(el, ctx) {
    const result: Record<string, unknown> = {};
    let pageSetUpPrCache: Record<string, unknown> | undefined;

    // Resolve shared strings from context (XlsxReadContext)
    const strings: string[] =
      ctx && "sharedStrings" in ctx ? (ctx as XlsxReadContext).sharedStrings : [];

    // Sheet properties
    const sheetPrEl = findChild(el, "sheetPr");
    if (sheetPrEl) {
      const sp: Record<string, unknown> = {};
      if (attr(sheetPrEl, "syncHorizontal") === "1") sp.syncHorizontal = true;
      if (attr(sheetPrEl, "syncVertical") === "1") sp.syncVertical = true;
      if (attr(sheetPrEl, "syncRef")) sp.syncRef = attr(sheetPrEl, "syncRef");
      if (attr(sheetPrEl, "transitionEvaluation") === "1") sp.transitionEvaluation = true;
      if (attr(sheetPrEl, "transitionEntry") === "1") sp.transitionEntry = true;
      if (attr(sheetPrEl, "published") === "1") sp.published = true;
      if (attr(sheetPrEl, "filterMode") === "1") sp.filterMode = true;
      if (attr(sheetPrEl, "enableFormatConditionsCalculation") === "1")
        sp.enableFormatConditionsCalculation = true;

      const outlinePr = findChild(sheetPrEl, "outlinePr");
      if (outlinePr) {
        if (attr(outlinePr, "applyStyles") === "1") sp.outlineApplyStyles = true;
        if (attr(outlinePr, "showOutlineSymbols") === "0") sp.outlineShowSymbols = false;
        if (attr(outlinePr, "summaryBelow") === "0") sp.outlineSummaryBelow = false;
        if (attr(outlinePr, "summaryRight") === "0") sp.outlineSummaryRight = false;
      }

      // pageSetUpPr (inside sheetPr) — stash on result.pageSetup; merged into
      // the <pageSetup> parse below, which owns result.pageSetup.
      const pageSetUpPr = findChild(sheetPrEl, "pageSetUpPr");
      if (pageSetUpPr) {
        const psup: Record<string, unknown> = {};
        if (attr(pageSetUpPr, "fitToPage") === "1") psup.fitToPage = true;
        if (attr(pageSetUpPr, "autoPageBreaks") === "1") psup.autoPageBreaks = true;
        if (Object.keys(psup).length > 0) pageSetUpPrCache = psup;
      }
      if (Object.keys(sp).length > 0) result.sheetPr = sp;

      // Tab color
      const tabColorEl = findChild(sheetPrEl, "tabColor");
      if (tabColorEl) {
        const tc: Record<string, unknown> = {};
        if (attr(tabColorEl, "rgb")) tc.rgb = attr(tabColorEl, "rgb");
        if (attrNum(tabColorEl, "theme") !== undefined) tc.theme = attrNum(tabColorEl, "theme");
        if (attrNum(tabColorEl, "tint") !== undefined) tc.tint = attrNum(tabColorEl, "tint");
        if (attrNum(tabColorEl, "indexed") !== undefined)
          tc.indexed = attrNum(tabColorEl, "indexed");
        result.tabColor = tc;
      }
    }

    // Sheet views
    const sheetViewsEl = findChild(el, "sheetViews");
    if (sheetViewsEl) {
      const svEl = findChild(sheetViewsEl, "sheetView");
      if (svEl) {
        const sv: Record<string, unknown> = {};
        if (attr(svEl, "showGridLines") === "0") sv.showGridLines = false;
        if (attr(svEl, "showRowColHeaders") === "0") sv.showRowColHeaders = false;
        if (attr(svEl, "showZeros") === "0") sv.showZeros = false;
        const zs = attrNum(svEl, "zoomScale");
        if (zs !== undefined) sv.zoomScale = zs;
        if (attr(svEl, "tabSelected") !== undefined)
          sv.tabSelected = attr(svEl, "tabSelected") !== "0";
        if (attr(svEl, "rightToLeft") === "1") sv.rightToLeft = true;
        if (attr(svEl, "windowProtection") === "1") sv.windowProtection = true;
        if (attr(svEl, "showFormulas") === "1") sv.showFormulas = true;
        if (attr(svEl, "showRuler") === "0") sv.showRuler = false;
        if (attr(svEl, "showOutlineSymbols") === "0") sv.showOutlineSymbols = false;
        if (attr(svEl, "defaultGridColor") === "0") sv.defaultGridColor = false;
        if (attr(svEl, "showWhiteSpace") === "0") sv.showWhiteSpace = false;
        if (attr(svEl, "view")) sv.view = attr(svEl, "view");
        const colorId = attrNum(svEl, "colorId");
        if (colorId !== undefined) sv.colorId = colorId;
        const zsn = attrNum(svEl, "zoomScaleNormal");
        if (zsn !== undefined) sv.zoomScaleNormal = zsn;
        const zssl = attrNum(svEl, "zoomScaleSheetLayoutView");
        if (zssl !== undefined) sv.zoomScaleSheetLayoutView = zssl;
        const zspl = attrNum(svEl, "zoomScalePageLayoutView");
        if (zspl !== undefined) sv.zoomScalePageLayoutView = zspl;
        result.sheetView = sv;

        // Freeze pane
        const paneEl = findChild(svEl, "pane");
        if (paneEl && attr(paneEl, "state") === "frozen") {
          const fp: Record<string, unknown> = {};
          const ys = attrNum(paneEl, "ySplit");
          if (ys && ys > 0) fp.row = ys;
          const xs = attrNum(paneEl, "xSplit");
          if (xs && xs > 0) fp.col = xs;
          if (Object.keys(fp).length > 0) result.freezePanes = fp;
        }
      }
    }

    // Sheet format properties
    const sfpEl = findChild(el, "sheetFormatPr");
    if (sfpEl) {
      const sfp: Record<string, unknown> = {};
      const bcw = attrNum(sfpEl, "baseColWidth");
      if (bcw !== undefined) sfp.baseColWidth = bcw;
      const dcw = attrNum(sfpEl, "defaultColWidth");
      if (dcw !== undefined) sfp.defaultColWidth = dcw;
      const drh = attrNum(sfpEl, "defaultRowHeight");
      if (drh !== undefined) sfp.defaultRowHeight = drh;
      if (attr(sfpEl, "zeroHeight") === "1") sfp.zeroHeight = true;
      if (attr(sfpEl, "thickTop") === "1") sfp.thickTop = true;
      if (attr(sfpEl, "thickBottom") === "1") sfp.thickBottom = true;
      const olr = attrNum(sfpEl, "outlineLevelRow");
      if (olr !== undefined) sfp.outlineLevelRow = olr;
      const olc = attrNum(sfpEl, "outlineLevelCol");
      if (olc !== undefined) sfp.outlineLevelCol = olc;
      result.sheetFormatPr = sfp;
    }

    // Columns
    const colsEl = findChild(el, "cols");
    if (colsEl) {
      const columns: Record<string, unknown>[] = [];
      for (const colEl of colsEl.elements ?? []) {
        if (colEl.name !== "col") continue;
        const col: Record<string, unknown> = {};
        col.min = attrNum(colEl, "min") ?? 0;
        col.max = attrNum(colEl, "max") ?? 0;
        const w = attrNum(colEl, "width");
        if (w !== undefined) col.width = w;
        if (attr(colEl, "hidden") === "1") col.hidden = true;
        if (attr(colEl, "customWidth") === "1") col.customWidth = true;
        const ol = attrNum(colEl, "outlineLevel");
        if (ol !== undefined) col.outlineLevel = ol;
        if (attr(colEl, "collapsed") === "1") col.collapsed = true;
        if (attr(colEl, "bestFit") === "1") col.bestFit = true;
        if (attr(colEl, "phonetic") === "1") col.phonetic = true;
        columns.push(col);
      }
      if (columns.length > 0) result.columns = columns;
    }

    // Sheet protection
    const protEl = findChild(el, "sheetProtection");
    if (protEl?.attributes) {
      const prot: Record<string, unknown> = {};
      if (attr(protEl, "password")) prot.password = attr(protEl, "password");
      if (attr(protEl, "algorithmName")) prot.algorithmName = attr(protEl, "algorithmName");
      if (attr(protEl, "hashValue")) prot.hashValue = attr(protEl, "hashValue");
      if (attr(protEl, "saltValue")) prot.saltValue = attr(protEl, "saltValue");
      if (attrNum(protEl, "spinCount") !== undefined) prot.spinCount = attrNum(protEl, "spinCount");
      if (attr(protEl, "sheet") === "1") prot.sheet = true;
      if (attr(protEl, "objects") === "1") prot.objects = true;
      if (attr(protEl, "scenarios") === "1") prot.scenarios = true;
      if (attr(protEl, "formatCells") === "0") prot.formatCells = false;
      if (attr(protEl, "formatColumns") === "0") prot.formatColumns = false;
      if (attr(protEl, "formatRows") === "0") prot.formatRows = false;
      if (attr(protEl, "insertColumns") === "0") prot.insertColumns = false;
      if (attr(protEl, "insertRows") === "0") prot.insertRows = false;
      if (attr(protEl, "insertHyperlinks") === "0") prot.insertHyperlinks = false;
      if (attr(protEl, "deleteColumns") === "0") prot.deleteColumns = false;
      if (attr(protEl, "deleteRows") === "0") prot.deleteRows = false;
      if (attr(protEl, "selectLockedCells") === "1") prot.selectLockedCells = true;
      if (attr(protEl, "sort") === "0") prot.sort = false;
      if (attr(protEl, "autoFilter") === "0") prot.autoFilter = false;
      if (attr(protEl, "pivotTables") === "0") prot.pivotTables = false;
      if (attr(protEl, "selectUnlockedCells") === "1") prot.selectUnlockedCells = true;
      result.protection = prot;
    }

    // Protected ranges
    const prEl = findChild(el, "protectedRanges");
    if (prEl) {
      const ranges: Record<string, unknown>[] = [];
      for (const rEl of prEl.elements ?? []) {
        if (rEl.name !== "protectedRange") continue;
        const r: Record<string, unknown> = {};
        r.sqref = attr(rEl, "sqref") ?? "";
        r.name = attr(rEl, "name") ?? "";
        if (attr(rEl, "password")) r.password = attr(rEl, "password");
        if (attr(rEl, "algorithmName")) r.algorithmName = attr(rEl, "algorithmName");
        if (attr(rEl, "hashValue")) r.hashValue = attr(rEl, "hashValue");
        if (attr(rEl, "saltValue")) r.saltValue = attr(rEl, "saltValue");
        if (attrNum(rEl, "spinCount") !== undefined) r.spinCount = attrNum(rEl, "spinCount");
        const sdEl = findChild(rEl, "securityDescriptor");
        if (sdEl) r.securityDescriptor = textOf(sdEl);
        ranges.push(r);
      }
      if (ranges.length > 0) result.protectedRanges = ranges;
    }

    // Auto filter
    const afEl = findChild(el, "autoFilter");
    if (afEl) {
      result.autoFilter = attr(afEl, "ref") ?? "";
    }

    // Merge cells
    const mcEl = findChild(el, "mergeCells");
    if (mcEl) {
      const merges: Record<string, unknown>[] = [];
      for (const mEl of mcEl.elements ?? []) {
        if (mEl.name !== "mergeCell") continue;
        const ref = attr(mEl, "ref") ?? "";
        const parts = ref.split(":");
        if (parts.length === 2) {
          const from = parseCellRef(parts[0]);
          const to = parseCellRef(parts[1]);
          if (from && to) merges.push({ from, to });
        }
      }
      if (merges.length > 0) result.mergeCells = merges;
    }

    // Conditional formatting
    const cfEls = el.elements?.filter((e) => e.name === "conditionalFormatting") ?? [];
    if (cfEls.length > 0) {
      const cfs: Record<string, unknown>[] = [];
      for (const cfEl of cfEls) {
        const sqref = attr(cfEl, "sqref") ?? "";
        const rules: Record<string, unknown>[] = [];
        for (const ruleEl of cfEl.elements ?? []) {
          if (ruleEl.name !== "cfRule") continue;
          const rule: Record<string, unknown> = {};
          rule.type = attr(ruleEl, "type");
          rule.priority = attrNum(ruleEl, "priority") ?? 1;
          if (attr(ruleEl, "operator")) rule.operator = attr(ruleEl, "operator");
          const dxfId = attrNum(ruleEl, "dxfId");
          if (dxfId !== undefined) rule.dxfId = dxfId;
          if (attr(ruleEl, "stopIfTrue") === "1") rule.stopIfTrue = true;
          if (attr(ruleEl, "timePeriod")) rule.timePeriod = attr(ruleEl, "timePeriod");
          const rank = attrNum(ruleEl, "rank");
          if (rank !== undefined) rule.rank = rank;
          if (attr(ruleEl, "equalAverage") === "1") rule.equalAverage = true;

          // Color scale
          const csEl = findChild(ruleEl, "colorScale");
          if (csEl) {
            const cfvo: Record<string, unknown>[] = [];
            const colors: string[] = [];
            for (const child of csEl.elements ?? []) {
              if (child.name === "cfvo") cfvo.push(parseCfvo(child));
              if (child.name === "color") {
                const rgb = attr(child, "rgb");
                if (rgb) colors.push(rgb.length === 8 ? rgb.slice(2) : rgb);
              }
            }
            rule.colorScale = { cfvo, colors };
          }

          // Data bar
          const dbEl = findChild(ruleEl, "dataBar");
          if (dbEl) {
            const cfvo: Record<string, unknown>[] = [];
            let color = "";
            for (const child of dbEl.elements ?? []) {
              if (child.name === "cfvo") cfvo.push(parseCfvo(child));
              if (child.name === "color") {
                const rgb = attr(child, "rgb");
                if (rgb) color = rgb.length === 8 ? rgb.slice(2) : rgb;
              }
            }
            rule.dataBar = { cfvo: cfvo as [any, any], color };
          }

          // Icon set
          const isEl = findChild(ruleEl, "iconSet");
          if (isEl) {
            const cfvo: Record<string, unknown>[] = [];
            for (const child of isEl.elements ?? []) {
              if (child.name === "cfvo") cfvo.push(parseCfvo(child));
            }
            const iconSet: Record<string, unknown> = { cfvo };
            if (attr(isEl, "iconSet")) iconSet.iconSet = attr(isEl, "iconSet");
            if (attr(isEl, "showValue") === "0") iconSet.showValue = false;
            if (attr(isEl, "percent") === "0") iconSet.percent = false;
            if (attr(isEl, "reverse") === "1") iconSet.reverse = true;
            rule.iconSet = iconSet;
          }

          // Formulas
          const formulas: string[] = [];
          for (const child of ruleEl.elements ?? []) {
            if (child.name === "formula") formulas.push(textOf(child) ?? "");
          }
          if (formulas.length > 0) rule.formulas = formulas;

          rules.push(rule);
        }
        cfs.push({ sqref, rules });
      }
      result.conditionalFormats = cfs;
    }

    // Data validations
    const dvEl = findChild(el, "dataValidations");
    if (dvEl) {
      const dvs: Record<string, unknown>[] = [];
      for (const dEl of dvEl.elements ?? []) {
        if (dEl.name !== "dataValidation") continue;
        const dv: Record<string, unknown> = {};
        dv.sqref = attr(dEl, "sqref") ?? "";
        if (attr(dEl, "type")) dv.type = attr(dEl, "type");
        if (attr(dEl, "operator")) dv.operator = attr(dEl, "operator");
        if (attr(dEl, "allowBlank") === "1") dv.allowBlank = true;
        if (attr(dEl, "showErrorMessage") === "1") dv.showErrorMessage = true;
        if (attr(dEl, "showInputMessage") === "1") dv.showInputMessage = true;
        if (attr(dEl, "errorTitle")) dv.errorTitle = attr(dEl, "errorTitle");
        if (attr(dEl, "error")) dv.error = attr(dEl, "error");
        if (attr(dEl, "promptTitle")) dv.promptTitle = attr(dEl, "promptTitle");
        if (attr(dEl, "prompt")) dv.prompt = attr(dEl, "prompt");
        if (attr(dEl, "errorStyle")) dv.errorStyle = attr(dEl, "errorStyle");
        if (attr(dEl, "imeMode")) dv.imeMode = attr(dEl, "imeMode");
        if (attr(dEl, "showDropDown") === "1") dv.showDropDown = true;

        const f1El = findChild(dEl, "formula1");
        if (f1El) dv.formula1 = textOf(f1El);
        const f2El = findChild(dEl, "formula2");
        if (f2El) dv.formula2 = textOf(f2El);

        dvs.push(dv);
      }
      result.dataValidations = dvs;
    }

    // Hyperlinks
    const hlEl = findChild(el, "hyperlinks");
    if (hlEl) {
      const hyperlinks: Record<string, unknown>[] = [];
      for (const hEl of hlEl.elements ?? []) {
        if (hEl.name !== "hyperlink") continue;
        const hl: Record<string, unknown> = {};
        hl.cell = attr(hEl, "ref") ?? "";
        const rId = hEl.attributes?.["r:id"] as string | undefined;
        const location = attr(hEl, "location");
        if (rId) hl.target = { type: "external", url: rId };
        else if (location) hl.target = { type: "internal", location };
        if (attr(hEl, "tooltip")) hl.tooltip = attr(hEl, "tooltip");
        if (attr(hEl, "display")) hl.display = attr(hEl, "display");
        hyperlinks.push(hl);
      }
      result.hyperlinks = hyperlinks;
    }

    // Print options
    const poEl = findChild(el, "printOptions");
    if (poEl) {
      const po: Record<string, unknown> = {};
      if (attr(poEl, "horizontalCentered") === "1") po.horizontalCentered = true;
      if (attr(poEl, "verticalCentered") === "1") po.verticalCentered = true;
      if (attr(poEl, "headings") === "1") po.headings = true;
      if (attr(poEl, "gridLines") === "1") po.gridLines = true;
      if (attr(poEl, "gridLinesSet") === "0") po.gridLinesSet = false;
      result.printOptions = po;
    }

    // Page setup
    const psEl = findChild(el, "pageSetup");
    if (psEl) {
      const ps: Record<string, unknown> = {};
      const pz = attrNum(psEl, "paperSize");
      if (pz !== undefined) ps.paperSize = pz;
      if (attr(psEl, "orientation")) ps.orientation = attr(psEl, "orientation");
      const sc = attrNum(psEl, "scale");
      if (sc !== undefined) ps.scale = sc;
      const ftw = attrNum(psEl, "fitToWidth");
      if (ftw !== undefined) ps.fitToWidth = ftw;
      const fth = attrNum(psEl, "fitToHeight");
      if (fth !== undefined) ps.fitToHeight = fth;
      if (attr(psEl, "pageOrder")) ps.pageOrder = attr(psEl, "pageOrder");
      if (attr(psEl, "useFirstPageNumber") === "1") ps.useFirstPageNumber = true;
      const fpn = attrNum(psEl, "firstPageNumber");
      if (fpn !== undefined) ps.firstPageNumber = fpn;
      if (pageSetUpPrCache) Object.assign(ps, pageSetUpPrCache);
      result.pageSetup = ps;
    } else if (pageSetUpPrCache) {
      result.pageSetup = pageSetUpPrCache;
    }

    // Header/footer
    const hfEl = findChild(el, "headerFooter");
    if (hfEl) {
      const hf: Record<string, unknown> = {};
      if (attr(hfEl, "differentOddEven") === "1") hf.differentOddEven = true;
      if (attr(hfEl, "differentFirst") === "1") hf.differentFirst = true;
      if (attr(hfEl, "scaleWithDoc") === "0") hf.scaleWithDoc = false;
      if (attr(hfEl, "alignWithMargins") === "0") hf.alignWithMargins = false;
      const oh = findChild(hfEl, "oddHeader");
      if (oh) hf.oddHeader = textOf(oh);
      const of2 = findChild(hfEl, "oddFooter");
      if (of2) hf.oddFooter = textOf(of2);
      const eh = findChild(hfEl, "evenHeader");
      if (eh) hf.evenHeader = textOf(eh);
      const ef = findChild(hfEl, "evenFooter");
      if (ef) hf.evenFooter = textOf(ef);
      const fh = findChild(hfEl, "firstHeader");
      if (fh) hf.firstHeader = textOf(fh);
      const ff = findChild(hfEl, "firstFooter");
      if (ff) hf.firstFooter = textOf(ff);
      result.headerFooter = hf;
    }

    // Ignored errors
    const ieEl = findChild(el, "ignoredErrors");
    if (ieEl) {
      const errors: Record<string, unknown>[] = [];
      for (const eEl of ieEl.elements ?? []) {
        if (eEl.name !== "ignoredError") continue;
        const ie: Record<string, unknown> = {};
        ie.sqref = attr(eEl, "sqref") ?? "";
        if (attr(eEl, "evalError") === "1") ie.evalError = true;
        if (attr(eEl, "twoDigitTextYear") === "1") ie.twoDigitTextYear = true;
        if (attr(eEl, "numberStoredAsText") === "1") ie.numberStoredAsText = true;
        if (attr(eEl, "formula") === "1") ie.formula = true;
        if (attr(eEl, "formulaRange") === "1") ie.formulaRange = true;
        if (attr(eEl, "unlockedFormula") === "1") ie.unlockedFormula = true;
        if (attr(eEl, "emptyCellReference") === "1") ie.emptyCellReference = true;
        if (attr(eEl, "listDataValidation") === "1") ie.listDataValidation = true;
        if (attr(eEl, "calculatedColumn") === "1") ie.calculatedColumn = true;
        errors.push(ie);
      }
      result.ignoredErrors = errors;
    }

    // Phonetic properties
    const ppEl = findChild(el, "phoneticPr");
    if (ppEl) {
      const pp: Record<string, unknown> = {};
      pp.fontId = attrNum(ppEl, "fontId") ?? 0;
      if (attr(ppEl, "type")) pp.type = attr(ppEl, "type");
      if (attr(ppEl, "alignment")) pp.alignment = attr(ppEl, "alignment");
      result.phoneticPr = pp;
    }

    // Sheet calc properties
    const scEl = findChild(el, "sheetCalcPr");
    if (scEl) {
      const sc: Record<string, unknown> = {};
      if (attr(scEl, "fullCalcOnLoad") === "1") sc.fullCalcOnLoad = true;
      result.sheetCalcPr = sc;
    }

    // Sheet data (rows and cells)
    const sheetDataEl = findChild(el, "sheetData");
    if (sheetDataEl) {
      const rows: Record<string, unknown>[] = [];
      for (const rowEl of sheetDataEl.elements ?? []) {
        if (rowEl.name !== "row") continue;
        const row: Record<string, unknown> = {};
        const rowNumber = attrNum(rowEl, "r");
        if (rowNumber !== undefined) row.rowNumber = rowNumber;
        const ht = attrNum(rowEl, "ht");
        if (ht !== undefined) row.height = ht;
        if (attr(rowEl, "hidden") === "1") row.hidden = true;
        if (attr(rowEl, "spans")) row.spans = attr(rowEl, "spans");
        if (attr(rowEl, "customFormat") === "1") row.customFormat = true;
        if (attr(rowEl, "thickTop") === "1") row.thickTop = true;
        if (attr(rowEl, "thickBot") === "1") row.thickBot = true;
        if (attr(rowEl, "ph") === "1") row.ph = true;

        const cells: Record<string, unknown>[] = [];
        for (const cellEl of rowEl.elements ?? []) {
          if (cellEl.name !== "c") continue;
          const cell: Record<string, unknown> = {};
          const ref = attr(cellEl, "r");
          if (ref) cell.reference = ref;
          const type = attr(cellEl, "t");
          const styleIdx = attrNum(cellEl, "s");
          if (styleIdx !== undefined) {
            // Resolve to a concrete StyleOptions so re-stringify registers it in
            // the fresh Styles table (whose indices may differ). Keep styleIndex
            // as a fallback when the styles table cannot be resolved.
            const resolved =
              ctx && "resolveStyle" in ctx
                ? (ctx as XlsxReadContext).resolveStyle(styleIdx)
                : undefined;
            if (resolved) {
              cell.style = resolved;
            } else {
              cell.styleIndex = styleIdx;
            }
          }

          // Cell value
          const vEl = findChild(cellEl, "v");
          const isEl = findChild(cellEl, "is");

          if (type === "s" && vEl) {
            // Shared string
            const idx = parseInt(textOf(vEl) ?? "", 10);
            cell.value = strings[idx] ?? "";
          } else if (type === "b" && vEl) {
            cell.value = textOf(vEl) === "1";
          } else if (type === "inlineStr" && isEl) {
            const t = findChild(isEl, "t");
            cell.value = textOf(t) ?? "";
          } else if (vEl) {
            const raw = textOf(vEl) ?? "";
            const num = Number(raw);
            cell.value = isNaN(num) ? raw : num;
          }

          // Formula
          const fEl = findChild(cellEl, "f");
          if (fEl) {
            const formula: Record<string, unknown> = { formula: textOf(fEl) ?? "" };
            const ft = attr(fEl, "t");
            if (ft && ft !== "normal") formula.type = ft;
            const fRef = attr(fEl, "ref");
            if (fRef) formula.reference = fRef;
            const fSi = attrNum(fEl, "si");
            if (fSi !== undefined) formula.sharedIndex = fSi;
            if (attr(fEl, "aca") === "1") formula.aca = true;
            if (attr(fEl, "ca") === "1") formula.ca = true;
            if (attr(fEl, "bx") === "1") formula.bx = true;
            cell.formula = formula as unknown as FormulaOptions;
          }

          cells.push(cell);
        }

        row.cells = cells;
        rows.push(row);
      }
      if (rows.length > 0) result.rows = rows;
    }

    return result as unknown as WorksheetOptions;
  },
};

// ── Stringify implementation ──

/**
 * Build the complete worksheet XML string.
 *
 * Zero-allocation fast path: directly concatenates XML string,
 * bypassing the IXmlableObject intermediate tree entirely.
 */
export function stringifyWorksheet(opts: WorksheetOptions, ctx: WorksheetContext): string {
  const sharedStrings = ctx.sharedStrings;
  const styles = ctx.styles;

  const rows = opts.rows ?? [];
  const columns = opts.columns ?? [];
  const mergeCells = opts.mergeCells ?? [];
  const protectedRanges = opts.protectedRanges ?? [];
  const ignoredErrors = opts.ignoredErrors ?? [];
  const rowBreaks = opts.rowBreaks ?? [];
  const colBreaks = opts.colBreaks ?? [];
  const customSheetViews = opts.customSheetViews ?? [];
  const cellWatches = opts.cellWatches ?? [];
  const controls = opts.controls ?? [];
  const customProperties = opts.customProperties ?? [];
  const oleObjects = opts.oleObjects ?? [];
  const webPublishItems = opts.webPublishItems ?? [];

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
  const hasTabColor = !!opts.tabColor;
  const hasOutline = columns.some((c) => c.outlineLevel !== undefined);
  const sp = opts.sheetPr;
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
  const hasPageSetUpPr =
    !!opts.pageSetup?.fitToWidth ||
    !!opts.pageSetup?.fitToHeight ||
    !!opts.pageSetup?.autoPageBreaks;
  if (hasTabColor || hasOutline || hasSheetPrAttrs || hasPageSetUpPr) {
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
    if (opts.tabColor) {
      const tc = opts.tabColor;
      const tcAttrs: Record<string, string | number | boolean | undefined> = {};
      if (tc.rgb) tcAttrs.rgb = tc.rgb;
      if (tc.theme !== undefined) tcAttrs.theme = tc.theme;
      if (tc.tint !== undefined) tcAttrs.tint = tc.tint;
      if (tc.indexed !== undefined) tcAttrs.indexed = tc.indexed;
      prParts.push(`<tabColor${attrs(tcAttrs)}/>`);
    }
    if (hasOutline) {
      const outAttrs: Record<string, string | number | boolean | undefined> = {
        summaryBelow: 1,
        summaryRight: 1,
      };
      if (sp?.outlineSummaryBelow === false) outAttrs.summaryBelow = 0;
      if (sp?.outlineSummaryRight === false) outAttrs.summaryRight = 0;
      if (sp?.outlineApplyStyles) outAttrs.applyStyles = 1;
      if (sp?.outlineShowSymbols === false) outAttrs.showOutlineSymbols = 0;
      prParts.push(`<outlinePr${attrs(outAttrs)}/>`);
    }
    // pageSetUpPr (inside sheetPr when fitToPage or autoPageBreaks needed)
    if (
      opts.pageSetup?.fitToWidth ||
      opts.pageSetup?.fitToHeight ||
      opts.pageSetup?.autoPageBreaks
    ) {
      const psupAttrs: Record<string, string | number | boolean | undefined> = {};
      if (opts.pageSetup?.fitToWidth || opts.pageSetup?.fitToHeight) psupAttrs.fitToPage = 1;
      if (opts.pageSetup?.autoPageBreaks) psupAttrs.autoPageBreaks = 1;
      prParts.push(`<pageSetUpPr${attrs(psupAttrs)}/>`);
    }
    const prAttrStr = Object.keys(prAttrs).length > 0 ? attrs(prAttrs) : "";
    p.push(`<sheetPr${prAttrStr}>${prParts.join("")}</sheetPr>`);
  }

  // Dimension — defines the used range of the sheet
  const maxRow = rows.length;
  let maxCol = 0;
  for (const row of rows) {
    if (row.cells && row.cells.length > maxCol) maxCol = row.cells.length;
  }
  if (maxRow > 0 && maxCol > 0) {
    const dimRef = `A1:${defaultCellRef(maxRow, maxCol)}`;
    p.push(`<dimension ref="${dimRef}"/>`);
  }

  // Sheet views
  const pivotSelXml = opts.sheetView?.pivotSelections
    ? opts.sheetView.pivotSelections.map((ps) => buildPivotSelectionXml(ps)).join("")
    : "";
  if (opts.freezePanes) {
    const fp = opts.freezePanes;
    const ySplit = fp.row ? fp.row : 0;
    const xSplit = fp.col ? fp.col : 0;
    const topRow = fp.row ? fp.row + 1 : 1;
    const leftCol = fp.col ? fp.col + 1 : 1;
    const topLeftCell = defaultCellRef(topRow, leftCol);
    const activePane =
      ySplit > 0 && xSplit > 0 ? "bottomRight" : ySplit > 0 ? "bottomLeft" : "topRight";
    const svAttrs = buildSheetViewAttrs(opts.sheetView);
    p.push(
      `<sheetViews><sheetView${svAttrs}>`,
      `<pane ySplit="${ySplit}" xSplit="${xSplit}" topLeftCell="${topLeftCell}" activePane="${activePane}" state="frozen"/>`,
      opts.selection ? buildSelectionXml(opts.selection) : "",
      pivotSelXml,
      "</sheetView></sheetViews>",
    );
  } else {
    const svAttrs = buildSheetViewAttrs(opts.sheetView);
    const innerXml = (opts.selection ? buildSelectionXml(opts.selection) : "") + pivotSelXml;
    if (innerXml) {
      p.push(`<sheetViews><sheetView${svAttrs}>${innerXml}</sheetView></sheetViews>`);
    } else {
      p.push(`<sheetViews><sheetView${svAttrs}/></sheetViews>`);
    }
  }

  // Sheet format — default row height
  if (opts.sheetFormatPr) {
    const sfp = opts.sheetFormatPr;
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
  if (columns.length > 0) {
    p.push("<cols>");
    for (const col of columns) {
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
  for (let i = 0; i < rows.length; i++) {
    const rowOpts = rows[i];
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
      p.push(`<row${attrsRaw(rowAttrs)}>`);
      for (let j = 0; j < rowOpts.cells.length; j++) {
        const cell = rowOpts.cells[j];
        const ref = cell.reference ?? defaultCellRef(rowNumber, j + 1);
        const cellStr = buildCellString(ref, cell, sharedStrings, styles);
        if (cellStr) p.push(cellStr);
      }
      p.push("</row>");
    } else {
      p.push(`<row${attrsRaw(rowAttrs)}/>`);
    }
  }
  p.push("</sheetData>");

  // Sheet calc properties (after sheetData per XSD sequence)
  if (opts.sheetCalcPr) {
    const scAttrs: string[] = [];
    if (opts.sheetCalcPr.fullCalcOnLoad) scAttrs.push('fullCalcOnLoad="1"');
    p.push(`<sheetCalcPr${scAttrs.length ? " " + scAttrs.join(" ") : ""}/>`);
  }

  // Row breaks (after sheetCalcPr per XSD sequence)
  if (rowBreaks.length > 0) {
    const brkParts = rowBreaks.map((b) => {
      const bAttrs: Record<string, string | number | boolean | undefined> = { id: b.id };
      if (b.min !== undefined) bAttrs.min = b.min;
      if (b.max !== undefined) bAttrs.max = b.max;
      if (b.manual) bAttrs.man = 1;
      if (b.pivot) bAttrs.pt = 1;
      return `<brk${attrs(bAttrs)}/>`;
    });
    p.push(
      `<rowBreaks count="${rowBreaks.length}" manualBreakCount="${rowBreaks.filter((b) => b.manual).length}">${brkParts.join("")}</rowBreaks>`,
    );
  }

  // Column breaks
  if (colBreaks.length > 0) {
    const brkParts = colBreaks.map((b) => {
      const bAttrs: Record<string, string | number | boolean | undefined> = { id: b.id };
      if (b.min !== undefined) bAttrs.min = b.min;
      if (b.max !== undefined) bAttrs.max = b.max;
      if (b.manual) bAttrs.man = 1;
      if (b.pivot) bAttrs.pt = 1;
      return `<brk${attrs(bAttrs)}/>`;
    });
    p.push(
      `<colBreaks count="${colBreaks.length}" manualBreakCount="${colBreaks.filter((b) => b.manual).length}">${brkParts.join("")}</colBreaks>`,
    );
  }

  // Custom properties (CT_CustomProperties, after colBreaks per XSD sequence)
  if (customProperties.length > 0) {
    const cpParts: string[] = ["<customProperties>"];
    for (const cp of customProperties) {
      cpParts.push(`<customPr name="${escapeXml(cp.name)}" r:id="${escapeXml(cp.rId)}"/>`);
    }
    cpParts.push("</customProperties>");
    p.push(cpParts.join(""));
  }

  // OLE size
  if (opts.oleSize) {
    p.push(`<oleSize ref="${escapeXml(opts.oleSize)}"/>`);
  }

  // Custom sheet views (after oleSize per XSD sequence)
  if (customSheetViews.length > 0) {
    p.push("<customSheetViews>");
    for (const csv of customSheetViews) {
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
  if (cellWatches.length > 0) {
    p.push("<cellWatches>");
    for (const cw of cellWatches) {
      p.push(`<cellWatch r="${escapeXml(cw.r)}"/>`);
    }
    p.push("</cellWatches>");
  }

  // Data consolidation
  if (opts.dataConsolidate) {
    const dc = opts.dataConsolidate;
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
  if (opts.protection) {
    const prot = opts.protection;
    const protAttrs: Record<string, string | number | boolean | undefined> = {};
    if (prot.password) protAttrs.password = hashPassword(prot.password);
    // Auto-derive modern hash when password provided without explicit hashValue
    let derived: ReturnType<typeof derivePasswordHash> | undefined;
    if (prot.password !== undefined && prot.hashValue === undefined) {
      derived = derivePasswordHash(prot.password);
    }
    protAttrs.algorithmName = prot.algorithmName ?? derived?.algorithmName;
    protAttrs.hashValue = prot.hashValue ?? derived?.hashValue;
    protAttrs.saltValue = prot.saltValue ?? derived?.saltValue;
    if (prot.spinCount !== undefined) protAttrs.spinCount = prot.spinCount;
    else if (derived) protAttrs.spinCount = derived.spinCount;
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
  if (protectedRanges.length > 0) {
    const prParts: string[] = ["<protectedRanges>"];
    for (const pr of protectedRanges) {
      const prAttrs: Record<string, string | number | boolean | undefined> = {
        name: pr.name,
        sqref: pr.sqref,
      };
      if (pr.password) prAttrs.password = hashPassword(pr.password);
      // Auto-derive modern hash when password provided without explicit hashValue
      let prDerived: ReturnType<typeof derivePasswordHash> | undefined;
      if (pr.password !== undefined && pr.hashValue === undefined) {
        prDerived = derivePasswordHash(pr.password);
      }
      prAttrs.algorithmName = pr.algorithmName ?? prDerived?.algorithmName;
      prAttrs.hashValue = pr.hashValue ?? prDerived?.hashValue;
      prAttrs.saltValue = pr.saltValue ?? prDerived?.saltValue;
      if (pr.spinCount !== undefined) prAttrs.spinCount = pr.spinCount;
      else if (prDerived) prAttrs.spinCount = prDerived.spinCount;
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
  if (opts.scenarios) {
    const scParts: string[] = ["<scenarios"];
    const scAttrs: Record<string, string | number> = {};
    if (opts.scenarios.current !== undefined) scAttrs.current = opts.scenarios.current;
    if (opts.scenarios.show !== undefined) scAttrs.show = opts.scenarios.show;
    scParts[0] = `<scenarios${attrs(scAttrs)}>`;

    for (const scenario of opts.scenarios.scenarios) {
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
        if (cell.undone) icAttrs.undone = true;
        sParts.push(`<inputCells${attrs(icAttrs)}/>`);
      }
      sParts.push("</scenario>");
      scParts.push(sParts.join(""));
    }
    scParts.push("</scenarios>");
    p.push(scParts.join(""));
  }

  // Auto filter
  if (opts.autoFilter) {
    if (typeof opts.autoFilter === "string") {
      p.push(selfCloseElement("autoFilter", attrs({ ref: opts.autoFilter })));
    } else {
      const af = opts.autoFilter;
      const inner: string[] = [];
      for (const t10 of af.top10 ?? []) {
        const fcAttrs: Record<string, string | number | boolean | undefined> = {
          colId: t10.colId,
        };
        if (t10.hiddenButton) fcAttrs.hiddenButton = 1;
        if (t10.showButton === false) fcAttrs.showButton = 0;
        const t10Attrs: Record<string, string | number | boolean | undefined> = { val: t10.val };
        if (t10.top === false) t10Attrs.top = 0;
        if (t10.percent) t10Attrs.percent = 1;
        if (t10.filterVal !== undefined) t10Attrs.filterVal = t10.filterVal;
        inner.push(`<filterColumn${attrs(fcAttrs)}><top10${attrs(t10Attrs)}/></filterColumn>`);
      }
      for (const cf of af.customFilters ?? []) {
        const fcAttrs: Record<string, string | number | boolean | undefined> = {
          colId: cf.colId,
        };
        if (cf.hiddenButton) fcAttrs.hiddenButton = 1;
        if (cf.showButton === false) fcAttrs.showButton = 0;
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
            `<filterColumn${attrs(fcAttrs)}><customFilters${attrs(cfAttrs)}>${filters.join("")}</customFilters></filterColumn>`,
          );
        }
      }
      // Simple filters (CT_Filters)
      for (const fi of af.filters ?? []) {
        const fcAttrs: Record<string, string | number | boolean | undefined> = {
          colId: fi.colId,
        };
        const filtersAttrs: Record<string, string | number | boolean | undefined> = {};
        if (fi.blank) filtersAttrs.blank = 1;
        if (fi.calendarType) filtersAttrs.calendarType = fi.calendarType;
        const valParts = (fi.values ?? []).map((v) => `<filter val="${escapeXml(v)}"/>`);
        inner.push(
          `<filterColumn${attrs(fcAttrs)}><filters${attrs(filtersAttrs)}>${valParts.join("")}</filters></filterColumn>`,
        );
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
        if (df.valIso !== undefined) dfAttrs.valIso = df.valIso;
        if (df.maxValIso !== undefined) dfAttrs.maxValIso = df.maxValIso;
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
  if (mergeCells.length > 0) {
    p.push(`<mergeCells count="${mergeCells.length}">`);
    for (const mc of mergeCells) {
      const fromRef = defaultCellRef(mc.from.row, mc.from.col);
      const toRef = defaultCellRef(mc.to.row, mc.to.col);
      p.push(selfCloseElement("mergeCell", attrs({ ref: `${fromRef}:${toRef}` })));
    }
    p.push("</mergeCells>");
  }

  // Phonetic properties (after mergeCells per XSD sequence)
  if (opts.phoneticPr) {
    const pp = opts.phoneticPr;
    const ppAttrs: Record<string, string | number> = { fontId: pp.fontId };
    if (pp.type && pp.type !== "fullwidthKatakana") ppAttrs.type = pp.type;
    if (pp.alignment && pp.alignment !== "left") ppAttrs.alignment = pp.alignment;
    p.push(selfCloseElement("phoneticPr", attrs(ppAttrs)));
  }

  // Conditional formatting
  const conditionalFormats = opts.conditionalFormats ?? [];
  if (conditionalFormats.length > 0) {
    for (const cf of conditionalFormats) {
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
            inner.push(buildCfvoXml(v));
          }
          for (const c of cs.colors) {
            inner.push(`<color rgb="FF${c}"/>`);
          }
          p.push(`<cfRule${attrs(ruleAttrs)}><colorScale>${inner.join("")}</colorScale></cfRule>`);
        }
        // Data bar
        else if (rule.type === "dataBar" && rule.dataBar) {
          const db = rule.dataBar;
          const inner: string[] = [];
          for (const v of db.cfvo) {
            inner.push(buildCfvoXml(v));
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
            inner.push(buildCfvoXml(v));
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
  const dataValidations = opts.dataValidations ?? [];
  if (dataValidations.length > 0) {
    const dvContainerAttrs: Record<string, string | number | boolean | undefined> = {
      count: dataValidations.length,
    };
    if (opts.dataValidationsDisablePrompts) dvContainerAttrs.disablePrompts = 1;
    p.push(`<dataValidations${attrs(dvContainerAttrs)}>`);
    for (const dv of dataValidations) {
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
  const hyperlinks = opts.hyperlinks ?? [];
  if (hyperlinks.length > 0) {
    p.push("<hyperlinks>");
    let hlIdx = 0;
    for (const hl of hyperlinks) {
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
  if (opts.printOptions) {
    const po = opts.printOptions;
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
  if (opts.pageSetup) {
    const ps = opts.pageSetup;
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
  if (opts.headerFooter) {
    const hf = opts.headerFooter;
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
  if (opts.drawingHF) {
    const dhf = opts.drawingHF;
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
  if (opts.legacyDrawingHF) {
    p.push(`<legacyDrawingHF r:id="${escapeXml(opts.legacyDrawingHF)}"/>`);
  }

  // Ignored errors (after headerFooter per XSD sequence)
  if (ignoredErrors.length > 0) {
    const ieParts: string[] = ["<ignoredErrors>"];
    for (const ie of ignoredErrors) {
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
  if (opts.backgroundImage) {
    p.push("<!--BACKGROUND_PICTURE-->");
  }

  // OLE objects (CT_OleObjects, after picture per XSD sequence)
  if (oleObjects.length > 0) {
    const oleParts: string[] = ["<oleObjects>"];
    for (const ole of oleObjects) {
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
  if (controls.length > 0) {
    const ctrlParts: string[] = ["<controls>"];
    for (const c of controls) {
      const cAttrs: string[] = [`shapeId="${c.shapeId}"`, `r:id="${escapeXml(c.rId)}"`];
      if (c.name) cAttrs.push(`name="${escapeXml(c.name)}"`);
      // controlPr (optional)
      const prAttrs: string[] = [];
      if (c.locked === false) prAttrs.push('locked="0"');
      if (c.uiObject) prAttrs.push('uiObject="1"');
      if (c.recalcAlways) prAttrs.push('recalcAlways="1"');
      if (c.linkedCell) prAttrs.push(`linkedCell="${escapeXml(c.linkedCell)}"`);
      if (c.listFillRange) prAttrs.push(`listFillRange="${escapeXml(c.listFillRange)}"`);
      if (c.cf) prAttrs.push(`cf="${escapeXml(c.cf)}"`);
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
  if (webPublishItems.length > 0) {
    const wpParts: string[] = [`<webPublishItems count="${webPublishItems.length}">`];
    for (const wpi of webPublishItems) {
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
  if (opts.ext) {
    p.push(`<extLst>${opts.ext}</extLst>`);
  }

  p.push("</worksheet>");
  return p.join("");
}

// ── Stringify helpers ──

function buildCfvoXml(cfvo: CfvoOptions): string {
  const a: Record<string, string | number | boolean | undefined> = { type: cfvo.type };
  if (cfvo.val !== undefined) a.val = cfvo.val;
  if (cfvo.gte === false) a.gte = 0;
  return `<cfvo${attrs(a)}/>`;
}

function buildSheetViewAttrs(sv?: SheetViewOptions): string {
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

function buildSelectionXml(sel: SelectionOptions): string {
  const selAttrs: Record<string, string | number | boolean | undefined> = {};
  if (sel.pane) selAttrs.pane = sel.pane;
  if (sel.activeCell) selAttrs.activeCell = sel.activeCell;
  if (sel.activeCellId !== undefined) selAttrs.activeCellId = sel.activeCellId;
  if (sel.sqref) selAttrs.sqref = sel.sqref;
  return `<selection${attrs(selAttrs)}/>`;
}

function buildPivotSelectionXml(_ps: PivotSelectionOptions): string {
  // pivotSelection is optional; omit if no meaningful pivotArea can be constructed.
  // An empty <pivotArea/> causes Excel to reject the file.
  return "";
}

function hashPassword(password: string): string {
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

function buildFormulaString(fOpts: FormulaOptions): string {
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

function buildCellString(
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
    const fStr = buildFormulaString(cell.formula);
    let vStr = "";
    if (value === null || value === undefined) {
      return `<c${attrsRaw(cellAttrs)}>${fStr}</c>`;
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
      vStr = `<v>${dateToSerialNumber(value)}</v>`;
    }
    if (vStr) {
      return `<c${attrsRaw(cellAttrs)}>${fStr}${vStr}</c>`;
    }
    return `<c${attrsRaw(cellAttrs)}>${fStr}</c>`;
  }

  if (value === null || value === undefined) {
    if (cell.styleIndex !== undefined) {
      return selfCloseElement("c", attrsRaw(cellAttrs));
    }
    return "";
  }

  // Rich text value (RichTextOptions)
  if (typeof value === "object" && !(value instanceof Date)) {
    if (sharedStrings) {
      cellAttrs.t = "s";
      const idx = sharedStrings.registerRich(value);
      return `<c${attrsRaw(cellAttrs)}><v>${idx}</v></c>`;
    }
    cellAttrs.t = "inlineStr";
    return `<c${attrsRaw(cellAttrs)}><is>${buildRstXml(value)}</is></c>`;
  }

  if (typeof value === "string") {
    if (sharedStrings) {
      cellAttrs.t = "s";
      const idx = sharedStrings.register(value);
      return `<c${attrsRaw(cellAttrs)}><v>${idx}</v></c>`;
    }
    cellAttrs.t = "inlineStr";
    return `<c${attrsRaw(cellAttrs)}><is><t>${escapeXml(value)}</t></is></c>`;
  }

  if (typeof value === "number") {
    return `<c${attrsRaw(cellAttrs)}><v>${value}</v></c>`;
  }

  if (typeof value === "boolean") {
    cellAttrs.t = "b";
    return `<c${attrsRaw(cellAttrs)}><v>${value ? 1 : 0}</v></c>`;
  }

  if (value instanceof Date) {
    const serial = dateToSerialNumber(value);
    return `<c${attrsRaw(cellAttrs)}><v>${serial}</v></c>`;
  }

  return "";
}

function defaultCellRef(row: number, col: number): string {
  return columnToLetter(col) + row;
}

function columnToLetter(col: number): string {
  let result = "";
  let n = col;
  while (n > 0) {
    const remainder = (n - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

function dateToSerialNumber(date: Date): number {
  const epoch = new Date(1899, 11, 30);
  const msPerDay = 86400000;
  return (date.getTime() - epoch.getTime()) / msPerDay;
}

// ── Parse helpers ──

function parseCfvo(el: XmlElement): Record<string, unknown> {
  const result: Record<string, unknown> = {};
  result.type = attr(el, "type") ?? "num";
  const val = attr(el, "val");
  if (val !== undefined) result.val = isNaN(Number(val)) ? val : Number(val);
  if (attr(el, "gte") === "0") result.gte = false;
  return result;
}

function parseCellRef(ref: string): { row: number; col: number } | undefined {
  const match = ref.match(/^([A-Z]+)(\d+)$/);
  if (!match) return undefined;
  const colStr = match[1];
  const row = parseInt(match[2], 10);
  let col = 0;
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }
  return { row, col };
}
