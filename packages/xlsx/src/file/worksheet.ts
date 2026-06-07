/**
 * Worksheet XML generation — pure functions for xl/worksheets/sheet{n}.xml.
 *
 * All interfaces and the zero-allocation string concatenation fast path
 * are preserved. The `Worksheet` class has been replaced by `buildWorksheetXml()`.
 *
 * @module
 */

import type { ChartSpaceOptions } from "@office-open/core";

import type { PivotTableOptions } from "./pivot";
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

// Re-exported from the descriptor for use by the compiler.
// The actual implementation lives in compile/descriptors/worksheet.ts.
export { stringifyWorksheet as buildWorksheetXml } from "../compile/descriptors/worksheet";
