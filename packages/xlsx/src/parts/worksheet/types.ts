/**
 * Worksheet — option types for xl/worksheets/sheet{n}.xml.
 *
 * @module
 */
import type {
  BasePictureOptions,
  ChartSpaceOptions,
  DataType,
  NonVisualDrawingPropertiesOptions,
  PositiveUniversalMeasure,
  UniversalMeasure,
} from "@office-open/core";
import type {
  BlackWhiteMode,
  BlipEffectsOptions,
  GraphicFrameLockingOptions,
  PictureLockingOptions,
  ShapePropertiesOptions,
  SourceRectangleOptions,
  TextHyperlinkOptions,
} from "@office-open/core/drawing";

import type {
  ConnectorOptions,
  DrawingAnchorOptions,
  GroupOptions,
  ShapeOptions,
} from "../drawing";
import type { PivotTableOptions } from "../pivot";
import type { PivotAreaOptions } from "../pivot/pivot-utils";
import type { QueryTableOptions } from "../query-table";
import type { SharedStrings } from "../shared-strings";
import type { Styles, StyleOptions } from "../styles";
import type { TableOptions } from "../table";
import type { SingleXmlCellOptions } from "../xml-mapping";

// ── Option interfaces ──

export interface ColumnOptions {
  min: number;
  max: number;
  width?: number;
  hidden?: boolean;
  customWidth?: boolean;
  outlineLevel?: number;
  collapsed?: boolean;
  /** Best-fit column width (CT_Col `@bestFit`) */
  bestFit?: boolean;
  /** Phonetic text for CJK (CT_Col `@phonetic`) */
  phonetic?: boolean;
}

export interface RowOptions {
  cells?: CellOptions[];
  height?: number | UniversalMeasure;
  hidden?: boolean;
  rowNumber?: number;
  /** Spans for the row, e.g. "1:15" (CT_Row `@spans`) */
  spans?: string;
  /** Custom format applied (CT_Row `@customFormat`) */
  customFormat?: boolean;
  /** Thick top border (CT_Row `@thickTop`) */
  thickTop?: boolean;
  /** Thick bottom border (CT_Row `@thickBot`) */
  thickBot?: boolean;
  /** Phonetic text (CT_Row `@ph`) */
  phonetic?: boolean;
  /** Row outline level for grouping (CT_Row `@outlineLevel`, 0-7) */
  outlineLevel?: number;
  /** Collapsed outline state (CT_Row `@collapsed`) */
  collapsed?: boolean;
  /**
   * Row style: either style options (resolved to an index at compile time) or,
   * as a round-trip fallback, the raw styles.xml cellXfs index (CT_Row `@s`).
   */
  style?: StyleOptions | number;
}

/** Rich text run properties (CT_RPrElt). */
export interface RichTextRunPropertiesOptions {
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
  properties?: RichTextRunPropertiesOptions;
  /** Run text content */
  text: string;
}

/** Phonetics run for CJK (CT_PhoneticRun → rPh). */
export interface PhoneticRunOptions {
  /** Start byte offset in base text (CT_PhoneticRun `@sb`) */
  startByte: number;
  /** End byte offset in base text (CT_PhoneticRun `@eb`) */
  endByte: number;
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
  /** Phonetic font/alignment settings (CT_PhoneticPr, si-level trailing element) */
  phoneticProperties?: PhoneticPropertiesOptions;
}

export interface CellOptions {
  value?: string | number | boolean | Date | RichTextOptions | null;
  reference?: string;
  /**
   * Cell style: either style options (resolved to an index at compile time)
   * or, as a round-trip fallback, the raw styles.xml cellXfs index carried
   * from a source workbook whose style table could not be resolved.
   */
  style?: StyleOptions | number;
  /**
   * Formula options. When set, value becomes the cached result. A bare string
   * is shorthand for `{ formula: "..." }`.
   */
  formula?: string | FormulaOptions;
  /** Cell metadata block id, 1-based index into metadata.cellMetadata (CT_Cell `@cm`) */
  cellMetadataId?: number;
  /** Value metadata block id, 1-based index into metadata.valueMetadata (CT_Cell `@vm`) */
  valueMetadataId?: number;
  /**
   * Error value (CT_Cell `t="e"`) — an Excel error literal like `#N/A` or
   * `#DIV/0!`. Emitted as `t="e"` instead of registering into the shared
   * string table, matching the source cell type.
   */
  error?: string;
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
  /** Always calculate array (CT_CellFormula `@aca`) */
  aca?: boolean;
  /** 2-D data table (CT_CellFormula `@dt2D`) */
  dt2D?: boolean;
  /** Data table row (CT_CellFormula `@dtr`) */
  dtr?: boolean;
  /** Delete input cell 1 (CT_CellFormula `@del1`) */
  del1?: boolean;
  /** Delete input cell 2 (CT_CellFormula `@del2`) */
  del2?: boolean;
  /** Input cell 1 reference (CT_CellFormula `@r1`) */
  inputCell1?: string;
  /** Input cell 2 reference (CT_CellFormula `@r2`) */
  inputCell2?: string;
  /** Calculate cell (CT_CellFormula `@ca`) */
  calculateCell?: boolean;
  /** Array formula context (CT_CellFormula `@bx`) */
  arrayContext?: boolean;
}

/** Input cell for a what-if scenario (maps to CT_InputCells). */
export interface ScenarioCellOptions {
  /** Cell reference, e.g. "B2" (CT_InputCells `@r`) */
  reference: string;
  /** Cell value for this scenario */
  val: string | number;
  /** Whether the value is deleted */
  deleted?: boolean;
  /** Whether undone (CT_InputCells `@undone`) */
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
  /** Merged range reference, e.g. "A1:D1" (CT_MergeCell `ref` attribute) */
  ref: string;
}

export interface SheetProtectionOptions {
  /**
   * Plain-text password — legacy Excel hash is computed automatically on
   * stringify. Authoring-only: not carried back by parse (the source stores a
   * hash; re-hashing it would corrupt it). Use the algorithmName quadruplet
   * below for round-trip.
   */
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
  /**
   * Plain-text password — legacy hash computed automatically on stringify.
   * Authoring-only: not carried back by parse (see SheetProtectionOptions).
   */
  password?: string;
  /** Modern encryption: algorithm name */
  algorithmName?: string;
  /** Modern encryption: base64-encoded hash value */
  hashValue?: string;
  /** Modern encryption: base64-encoded salt value */
  saltValue?: string;
  /** Modern encryption: spin count */
  spinCount?: number;
  /** Security descriptor (SID string, emitted as the attribute form) */
  securityDescriptor?: string;
}

export interface FreezePaneOptions {
  /** Row split position (1-based, freezes rows above) */
  row?: number;
  /** Column split position (1-based, freezes columns to the left) */
  col?: number;
  /** Split panes without freezing (CT_Pane `@state="split"`); default is frozen */
  split?: boolean;
  /**
   * Top-left cell of the scrollable pane (CT_Pane `@topLeftCell`); defaults to
   * the cell just past the split. Round-trip keeps a scrolled position.
   */
  topLeftCell?: string;
  /** Active pane (CT_Pane `@activePane`); defaults to the pane past the split */
  activePane?: "bottomRight" | "topRight" | "bottomLeft" | "topLeft";
}

/**
 * Picture anchored to a worksheet cell.
 *
 * Extends the cross-format {@link BasePictureOptions} (binary data + non-visual
 * drawing properties) with the full spreadsheet-drawing anchor
 * {@link DrawingAnchorOptions} (1-based from/to cell corners, anchor type,
 * extent). The base cNvPr fields (name/description/title/hidden) flow through
 * to the drawing's cNvPr.
 */
export interface PictureOptions extends Omit<BasePictureOptions, "type">, DrawingAnchorOptions {
  type: "png" | "jpg" | "wmf" | "emf";
  /** Round-tripped pic/spPr (rotation/flip/bwMode/fill beyond position). */
  spPr?: ShapePropertiesOptions;
  /** Blip crop (a:srcRect); an empty object round-trips the bare marker. */
  sourceRectangle?: SourceRectangleOptions;
  /** Black/white mode (spPr/@bwMode); absent = attribute omitted. */
  blackWhiteMode?: BlackWhiteMode;
  /**
   * Relative-resize hint (cNvPicPr/@preferRelativeResize). Absent = attribute
   * omitted (defaults true); explicit true/false round-trips the attribute.
   */
  preferRelativeResize?: boolean;
  /** Image adjustment effects carried inside a:blip (a:lum, a:duotone, …). */
  blipEffects?: BlipEffectsOptions;
  /** Picture locks (cNvPicPr/a:picLocks); absent = empty cNvPicPr. */
  locking?: PictureLockingOptions;
  /**
   * Click hyperlink on the picture itself (a:hlinkClick inside xdr:cNvPr) —
   * jump to a URL when the picture is clicked.
   */
  hyperlink?: TextHyperlinkOptions;
  /**
   * Document-order position inside the drawing part (z-order when objects
   * overlap). Round-trip only; fresh authoring uses insertion order.
   */
  zOrder?: number;
  /** Original cNvPr id (round-trip only). */
  shapeId?: number;
}

/**
 * Chart anchored to a worksheet. `title` stays the chart title (c:title);
 * the graphicFrame's cNvPr `@title` is too weak to justify a name clash, so
 * only name/description/hidden flow through to the drawing's cNvPr.
 */
export interface WorksheetChartOptions
  extends
    ChartSpaceOptions,
    DrawingAnchorOptions,
    Omit<NonVisualDrawingPropertiesOptions, "title"> {
  /** Frame locks (cNvGraphicFramePr/a:graphicFrameLocks); absent = empty. */
  frameLocks?: GraphicFrameLockingOptions;
  /**
   * Click hyperlink on the chart frame itself (a:hlinkClick inside xdr:cNvPr).
   */
  hyperlink?: TextHyperlinkOptions;
  /** Document-order position inside the drawing part (round-trip only). */
  zOrder?: number;
  /** Original cNvPr id (round-trip only). */
  shapeId?: number;
  /** Macro reference (CT_GraphicFrame/@macro); empty string round-trips. */
  macro?: string;
}

export interface SheetViewOptions {
  showGridLines?: boolean;
  showRowColHeaders?: boolean;
  showZeros?: boolean;
  zoomScale?: number;
  tabSelected?: boolean;
  rightToLeft?: boolean;
  /** Window protection (CT_SheetView `@windowProtection`) */
  windowProtection?: boolean;
  /** Show formulas instead of values (CT_SheetView `@showFormulas`) */
  showFormulas?: boolean;
  /** Show ruler (CT_SheetView `@showRuler`) */
  showRuler?: boolean;
  /** Show outline symbols (CT_SheetView `@showOutlineSymbols`) */
  showOutlineSymbols?: boolean;
  /** Default grid color (CT_SheetView `@defaultGridColor`) */
  defaultGridColor?: boolean;
  /** Show white space (CT_SheetView `@showWhiteSpace`) */
  showWhiteSpace?: boolean;
  /** View type (CT_SheetView `@view`) */
  view?: "normal" | "pageBreakPreview" | "pageLayout";
  /** Top-left visible cell (CT_SheetView `@topLeftCell`); round-trip only meaningful when scrolled */
  topLeftCell?: string;
  /** Tab color ID (CT_SheetView `@colorId`) */
  colorId?: number;
  /** Zoom scale for normal view (CT_SheetView `@zoomScaleNormal`) */
  zoomScaleNormal?: number;
  /** Zoom scale for sheet layout view (CT_SheetView `@zoomScaleSheetLayoutView`) */
  zoomScaleSheetLayoutView?: number;
  /** Zoom scale for page layout view (CT_SheetView `@zoomScalePageLayoutView`) */
  zoomScalePageLayoutView?: number;
}

export interface HyperlinkOptions {
  /** Cell reference, e.g. "A1" */
  cell: string;
  /** External target URL (CT_Hyperlink `@r:id`) */
  url?: string;
  /** Internal target, e.g. "Data!A1" (CT_Hyperlink `@location`; independent of url) */
  location?: string;
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
  /** Scale header/footer with document (CT_HeaderFooter `@scaleWithDoc`) */
  scaleWithDoc?: boolean;
  /** Align with page margins (CT_HeaderFooter `@alignWithMargins`) */
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
  /** Paper height (CT_PageSetup `@paperHeight`; a bare number means millimeters) */
  paperHeight?: number | PositiveUniversalMeasure;
  /** Paper width (CT_PageSetup `@paperWidth`; a bare number means millimeters) */
  paperWidth?: number | PositiveUniversalMeasure;
  /** Use printer defaults (CT_PageSetup `@usePrinterDefaults`, XSD default true — only false is emitted) */
  usePrinterDefaults?: boolean;
  /** Black and white printing (CT_PageSetup `@blackAndWhite`) */
  blackAndWhite?: boolean;
  /** Draft quality printing (CT_PageSetup `@draft`) */
  draft?: boolean;
  /** Print cell comments mode (CT_PageSetup `@cellComments`) */
  cellComments?: "none" | "asDisplayed" | "atEnd";
  /** Print error display mode (CT_PageSetup `@errors`) */
  errors?: "displayed" | "blank" | "dash" | "NA";
  /** Horizontal print DPI (CT_PageSetup `@horizontalDpi`) */
  horizontalDpi?: number;
  /** Vertical print DPI (CT_PageSetup `@verticalDpi`) */
  verticalDpi?: number;
  /** Number of copies to print (CT_PageSetup `@copies`) */
  copies?: number;
  /** Auto page breaks (CT_PageSetUpPr `@autoPageBreaks`) */
  autoPageBreaks?: boolean;
  /** Fit to page (CT_PageSetUpPr `@fitToPage`) */
  fitToPage?: boolean;
  /**
   * Relationship id to the printer settings binary part (CT_PageSetup
   * `@r:id`). Round-trip only: the referenced .bin part is not re-emitted,
   * so the id is not resolvable in a freshly generated workbook.
   */
  printerSettingsRId?: string;
}

export interface TabColorOptions {
  /** RGB color string, e.g. "FF0000" */
  rgb?: string;
  /** Theme color index (0-based) */
  theme?: number;
  /** Tint value (-1.0 to 1.0) */
  tint?: number;
  /** Indexed color (CT_Color `@indexed`) */
  indexed?: number;
}

/** Cell corner marker (CT_Marker): 0-based column/row plus EMU offsets. */
/**
 * Note anchor corner (CT_Marker). Mirrors the XML verbatim: 0-based col/row,
 * offsets in EMU — unlike DrawingAnchorOptions, which is 1-based for
 * authoring convenience. Same concept, two bases; the XML is 0-based either way.
 */
export interface AnchorMarkerOptions {
  /** 0-based column index */
  col: number;
  /** Offset within the column, in EMU (default: 0) */
  colOff?: number;
  /** 0-based row index */
  row: number;
  /** Offset within the row, in EMU (default: 0) */
  rowOff?: number;
}

/**
 * Object anchor (CT_ObjectAnchor) — the anchor inside commentPr, objectPr,
 * and controlPr alike. Mirrors the XML verbatim: 0-based markers
 * (CT_Marker), offsets in EMU.
 */
export interface ObjectAnchorOptions {
  /** Move with cells (default: false) */
  moveWithCells?: boolean;
  /** Size with cells (default: false) */
  sizeWithCells?: boolean;
  /** Anchor start corner (xdr:from, CT_Marker) */
  from: AnchorMarkerOptions;
  /** Anchor end corner (xdr:to, CT_Marker) */
  to: AnchorMarkerOptions;
}

/**
 * Comment property (CT_CommentPr).
 *
 * Parsed but never re-emitted: a commentPr alongside the sheet's legacy VML
 * note drawing produces a file Excel refuses to open — commentPr and the VML
 * note shape are rival property systems for the same note, and Excel reads
 * note properties from the shape's x:ClientData. The fields survive parse so
 * callers can inspect them; stringify always drops them.
 */
export interface CommentPropertiesOptions {
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
  /** Comment properties (CT_CommentPr) — parsed but never re-emitted (see CommentPropertiesOptions) */
  commentPr?: CommentPropertiesOptions;
  /**
   * Note shape anchor (x:Anchor in the VML part): from/to cell corners,
   * 0-based. Absent → the default 2×2-cell offset anchored at the comment's
   * cell.
   */
  anchor?: NoteAnchorOptions;
  /** Whether the note is pinned visible. Default false (hidden until hover). */
  visible?: boolean;
  /** Note shape size in points. Absent → 108 × 59.25 pt. */
  size?: { width: number; height: number };
}

/**
 * Note shape anchor (VML x:Anchor): two cell corners with offsets.
 *
 * Offsets are EMU in the public API (converted to/from the pixel values the
 * VML part stores at 96 DPI — exact both ways since px × 9525 = EMU).
 */
export interface NoteAnchorOptions {
  /** Anchor start corner */
  from: AnchorMarkerOptions;
  /** Anchor end corner */
  to: AnchorMarkerOptions;
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
  /** Error style (CT_DataValidation `@errorStyle`) */
  errorStyle?: "stop" | "warning" | "information";
  /** IME mode (CT_DataValidation `@imeMode`) */
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
  /** Show drop-down (CT_DataValidation `@showDropDown` — note inverted semantics in OOXML) */
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

/**
 * One CT_Color channel set for conditional-formatting colors. A color picks
 * exactly one channel (rgb, theme, or indexed); tint qualifies a theme slot.
 */
export interface CfColorOptions {
  /** RGB hex without alpha, e.g. "FF0000" */
  rgb?: string;
  /** Theme palette slot (CT_Color @theme) */
  theme?: number;
  /** Tint applied to the theme slot */
  tint?: number;
  /** Legacy palette index */
  indexed?: number;
}

/** Color scale rule configuration */
export interface ColorScaleOptions {
  /** Conditional format values (minimum 2, typically 2 or 3) */
  cfvo: CfvoOptions[];
  /** Colors for each value (same count as cfvo) */
  colors: CfColorOptions[];
}

/** Data bar rule configuration */
export interface DataBarOptions {
  /** Minimum and maximum value objects (exactly 2) */
  cfvo: [CfvoOptions, CfvoOptions];
  /** Bar color */
  color: CfColorOptions;
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
  /** Stop if true — skip remaining rules (CT_CfRule `@stopIfTrue`) */
  stopIfTrue?: boolean;
  /** Time period for date-based highlighting (CT_CfRule `@timePeriod`) */
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
  /** Rank for top/bottom rules (CT_CfRule `@rank`) */
  rank?: number;
  /** Bottom N instead of top N for top10 rules (CT_CfRule `@bottom`) */
  bottom?: boolean;
  /** Percent instead of item count for top10 rules (CT_CfRule `@percent`) */
  percent?: boolean;
  /** Search text for containsText/notContains/beginsWith/endsWith rules (CT_CfRule `@text`) */
  text?: string;
  /** Equal average flag (CT_CfRule `@equalAverage`) */
  equalAverage?: boolean;
  /** Above average flag (CT_CfRule `@aboveAverage`, default true for aboveAverage rules) */
  aboveAverage?: boolean;
  /** Standard deviations for above-average rules (CT_CfRule `@stdDev`) */
  stdDev?: number;
}

export interface ConditionalFormatOptions {
  /** Cell range, e.g. "A1:A10" */
  sqref: string;
  rules: ConditionalFormatRule[];
}

export interface Top10FilterOptions {
  top?: boolean;
  percent?: boolean;
  val: number;
  /** Filter value (CT_Top10 `@filterVal`) */
  filterVal?: number;
}

/** Custom filter comparison operator (ST_FilterOperator). */
export type CustomFilterOperator =
  | "equal"
  | "notEqual"
  | "greaterThan"
  | "greaterThanOrEqual"
  | "lessThan"
  | "lessThanOrEqual";

/** A single custom filter entry (CT_CustomFilter). */
export interface CustomFilterEntry {
  operator?: CustomFilterOperator;
  val: string;
}

/** Custom filters container (CT_CustomFilters): @and + customFilter entries. */
export interface CustomFiltersOptions {
  /** AND the entries instead of OR (CT_CustomFilters `@and`) */
  and?: boolean;
  entries: CustomFilterEntry[];
}

export interface SortCondition {
  /** Cell reference for the sort column, e.g. "B1" */
  ref: string;
  descending?: boolean;
  /** Sort by (CT_SortCondition `@sortBy`) */
  sortBy?: "value" | "cellColor" | "fontColor" | "icon";
  /** Custom sort list (CT_SortCondition `@customList`) */
  customList?: string;
  /** Differential format id for cell/font-color sorts (CT_SortCondition `@dxfId`) */
  dxfId?: number;
  /** Icon set for icon sorts (CT_SortCondition `@iconSet`, default "3Arrows") */
  iconSet?: IconSetType;
  /** Icon index within the set (CT_SortCondition `@iconId`) */
  iconId?: number;
}

/** Sort state (CT_SortState) — sort conditions nested in their XSD container. */
export interface SortStateOptions {
  /** Sort range, e.g. "A1:D10" (CT_SortState `@ref`, required) */
  ref: string;
  /** Column sort mode (CT_SortState `@columnSort`) */
  columnSort?: boolean;
  /** Case sensitive sorting (CT_SortState `@caseSensitive`) */
  caseSensitive?: boolean;
  /** Sort method (CT_SortState `@sortMethod`) */
  sortMethod?: "pinYin" | "stroke" | "none";
  /** Sort conditions, one per sort level (CT_SortState `sortCondition` children) */
  conditions: SortCondition[];
}

export interface AutoFilterOptions {
  /** Range, e.g. "A1:D10" */
  ref: string;
  /** Filter columns, one per filtered column (CT_FilterColumn) */
  columns?: FilterColumnOptions[];
  /** Sort state (CT_SortState child; `ref` typically spans the filter range) */
  sortState?: SortStateOptions;
}

/**
 * Filter column (CT_FilterColumn): column id plus exactly one filter variant
 * (the XSD content model is a choice among the six filter kinds below).
 */
export interface FilterColumnOptions {
  /** Column ID (CT_FilterColumn `@colId`) */
  colId: number;
  /** Hide auto-filter button (CT_FilterColumn `@hiddenButton`) */
  hiddenButton?: boolean;
  /** Show filter button (CT_FilterColumn `@showButton`) */
  showButton?: boolean;
  /** Top-N filter (CT_Top10) */
  top10?: Top10FilterOptions;
  /** Custom comparison filters (CT_CustomFilters) */
  customFilters?: CustomFiltersOptions;
  /** Value / date filters (CT_Filters) */
  filters?: FilterItemsOptions;
  /** Color filter (CT_ColorFilter) */
  colorFilter?: ColorFilterOptions;
  /** Icon filter (CT_IconFilter) */
  iconFilter?: IconFilterOptions;
  /** Dynamic filter (CT_DynamicFilter) */
  dynamicFilter?: DynamicFilterOptions;
}

/** Color filter (CT_ColorFilter) */
export interface ColorFilterOptions {
  /** Differential format id (used when no direct color is set) */
  dxfId?: number;
  /** Filter by cell color (CT_ColorFilter `@cellColor`) */
  cellColor?: boolean;
}

/** Icon filter (CT_IconFilter) */
export interface IconFilterOptions {
  /** Icon set name (CT_IconFilter `@iconSet`, ST_IconSetType — required) */
  iconSet: IconSetType;
  /** Icon ID within set (CT_IconFilter `@iconId`) */
  iconId?: number;
}

/** Filter items (CT_Filters) */
export interface FilterItemsOptions {
  /** Blank filter (CT_Filters `@blank`) */
  blank?: boolean;
  /** Calendar type (CT_Filters `@calendarType`, ST_CalendarType) */
  calendarType?:
    | "gregorian"
    | "gregorianUs"
    | "gregorianMeFrench"
    | "gregorianArabic"
    | "hijri"
    | "hebrew"
    | "taiwan"
    | "japan"
    | "thai"
    | "korea"
    | "saka"
    | "gregorianXlitEnglish"
    | "gregorianXlitFrench"
    | "none";
  /** Filter values (CT_Filters `filter` children) */
  values?: string[];
  /** Date group filters (CT_Filters `dateGroupItem` children) */
  dateGroupItems?: DateGroupFilterOptions[];
}

/** Dynamic filter (CT_DynamicFilter) */
export interface DynamicFilterOptions {
  /** Dynamic filter type (CT_DynamicFilter `@type`) */
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
  /** Filter value (CT_DynamicFilter `@val`) */
  val?: number;
  /** Maximum value (CT_DynamicFilter `@maxVal`) */
  maxVal?: number;
  /** Value ISO date string (CT_DynamicFilter `@valIso`) */
  valIso?: string;
  /** Max value ISO date string (CT_DynamicFilter `@maxValIso`) */
  maxValIso?: string;
}

/** Date group filter item (CT_DateGroupItem) */
export interface DateGroupFilterOptions {
  /** Date grouping level (CT_DateGroupItem `@dateTimeGrouping`) */
  dateTimeGrouping: "year" | "month" | "day" | "hour" | "minute" | "second";
  /** Year (CT_DateGroupItem `@year`) */
  year?: number;
  /** Month (1-12, CT_DateGroupItem `@month`) */
  month?: number;
  /** Day (1-31, CT_DateGroupItem `@day`) */
  day?: number;
  /** Hour (0-23, CT_DateGroupItem `@hour`) */
  hour?: number;
  /** Minute (0-59, CT_DateGroupItem `@minute`) */
  minute?: number;
  /** Second (0-59, CT_DateGroupItem `@second`) */
  second?: number;
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
export interface SheetFormatPropertiesOptions {
  /** Base column width (CT_SheetFormatPr `@baseColWidth`) */
  baseColWidth?: number;
  /** Default column width (CT_SheetFormatPr `@defaultColWidth`) */
  defaultColWidth?: number;
  /** Default row height */
  defaultRowHeight?: number;
  /** Zero height rows hidden (CT_SheetFormatPr `@zeroHeight`) */
  zeroHeight?: boolean;
  /** Thick top borders (CT_SheetFormatPr `@thickTop`) */
  thickTop?: boolean;
  /** Thick bottom borders (CT_SheetFormatPr `@thickBottom`) */
  thickBottom?: boolean;
  /** Outline level row (CT_SheetFormatPr `@outlineLevelRow`) */
  outlineLevelRow?: number;
  /** Outline level column (CT_SheetFormatPr `@outlineLevelCol`) */
  outlineLevelCol?: number;
}

/** Sheet properties extended options (CT_SheetPr attributes) */
export interface SheetPropertiesOptions {
  /** VBA code name (CT_SheetPr `@codeName`) */
  codeName?: string;
  /** Sync horizontal scroll (CT_SheetPr `@syncHorizontal`) */
  syncHorizontal?: boolean;
  /** Sync vertical scroll (CT_SheetPr `@syncVertical`) */
  syncVertical?: boolean;
  /** Sync reference (CT_SheetPr `@syncRef`) */
  syncRef?: string;
  /** Transition evaluation mode (CT_SheetPr `@transitionEvaluation`) */
  transitionEvaluation?: boolean;
  /** Transition entry mode (CT_SheetPr `@transitionEntry`) */
  transitionEntry?: boolean;
  /** Published to server (CT_SheetPr `@published`, XSD default true — only false is emitted) */
  published?: boolean;
  /** Filter mode (CT_SheetPr `@filterMode`) */
  filterMode?: boolean;
  /** Enable format conditions calculation (CT_SheetPr `@enableFormatConditionsCalculation`, XSD default true — only false is emitted) */
  enableFormatConditionsCalculation?: boolean;
  /** Outline apply styles (CT_OutlinePr `@applyStyles`) */
  outlineApplyStyles?: boolean;
  /** Outline show symbols (CT_OutlinePr `@showOutlineSymbols`) */
  outlineShowSymbols?: boolean;
  /** Outline summary rows below detail (CT_OutlinePr `@summaryBelow`) */
  outlineSummaryBelow?: boolean;
  /** Outline summary columns right of detail (CT_OutlinePr `@summaryRight`) */
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
export interface PhoneticPropertiesOptions {
  /** Font ID from the styles table (required) */
  fontId: number;
  /** Phonetic type (default: "fullwidthKatakana") */
  type?: "fullwidthKatakana" | "halfwidthKatakana" | "Hiragana" | "noConversion";
  /** Alignment (default: "left") */
  alignment?: "noControl" | "left" | "center" | "distributed";
}

/** Background image for a worksheet */
export interface SheetBackgroundImageOptions {
  data: DataType;
  type: "png" | "jpg";
}

/** Page break entry (CT_Break) */
export interface PageBreakOptions {
  /** Row or column ID (1-based) */
  id: number;
  /** Min value (CT_Break `@min`) */
  min?: number;
  /** Max value (CT_Break `@max`) */
  max?: number;
  /** Manual break (CT_Break `@man`) */
  manual?: boolean;
  /** Pivot break (CT_Break `@pt`) */
  pivot?: boolean;
}

/** Selection in sheet view (CT_Selection) */
export interface SelectionOptions {
  /** Pane (CT_Selection `@pane`) */
  pane?: "bottomRight" | "topRight" | "bottomLeft" | "topLeft";
  /** Active cell (CT_Selection `@activeCell`) */
  activeCell?: string;
  /** Active cell index (CT_Selection `@activeCellId`) */
  activeCellId?: number;
  /** Selected range (CT_Selection `@sqref`) */
  sqref?: string;
}

/** Pivot selection in sheet view (CT_PivotSelection) */
export interface PivotSelectionOptions {
  /** Pane (CT_PivotSelection `@pane`, default "topLeft") */
  pane?: "bottomRight" | "topRight" | "bottomLeft" | "topLeft";
  /** Show header (default false) */
  showHeader?: boolean;
  /** Label selection (default false) */
  label?: boolean;
  /** Data selection (default false) */
  data?: boolean;
  /** Extendable (default false) */
  extendable?: boolean;
  /** Selection count (default 0) */
  count?: number;
  /** Axis */
  axis?: "axisRow" | "axisCol" | "axisPage" | "axisValues";
  /** Dimension (default 0) */
  dimension?: number;
  /** Start (default 0) */
  start?: number;
  /** Min (default 0) */
  min?: number;
  /** Max (default 0) */
  max?: number;
  /** Active row (default 0) */
  activeRow?: number;
  /** Active column (default 0) */
  activeCol?: number;
  /** Previous row (default 0) */
  previousRow?: number;
  /** Previous column (default 0) */
  previousCol?: number;
  /** Click count (default 0) */
  click?: number;
  /**
   * Relationship id (r:id, OLAP pivot selection part). Round-trip only: the
   * referenced part is not re-emitted, so the id is not resolvable in a
   * freshly generated workbook.
   */
  rId?: string;
  /** Pivot area (required child) */
  pivotArea?: PivotAreaOptions;
}

/** Smart tag property (CT_CellSmartTagPr) */
export interface CellSmartTagPropertyOptions {
  /** Property key (required) */
  key: string;
  /** Property value (required) */
  val: string;
}

/** Smart tag on a cell (CT_CellSmartTag) */
export interface CellSmartTagOptions {
  /** Smart tag type index (required) */
  type: number;
  /** Deleted (default false) */
  deleted?: boolean;
  /** XML-based (default false) */
  xmlBased?: boolean;
  /** Properties */
  properties?: CellSmartTagPropertyOptions[];
}

/** Cell smart tags (CT_CellSmartTags) */
export interface CellSmartTagsOptions {
  /** Cell reference (required) */
  reference: string;
  /** Smart tags */
  smartTags: CellSmartTagOptions[];
}

/** Custom sheet view (CT_CustomSheetView) */
export interface CustomSheetViewOptions {
  /** GUID identifier (required, CT_CustomSheetView `@guid`) */
  guid: string;
  /** Zoom scale (CT_CustomSheetView `@scale`) */
  scale?: number;
  /** Show page breaks (CT_CustomSheetView `@showPageBreaks`) */
  showPageBreaks?: boolean;
  /** Show formulas (CT_CustomSheetView `@showFormulas`) */
  showFormulas?: boolean;
  /** Show grid lines (CT_CustomSheetView `@showGridLines`) */
  showGridLines?: boolean;
  /** Show row/column headers (CT_CustomSheetView `@showRowCol`) */
  showRowColHeaders?: boolean;
  /** Show outline symbols (CT_CustomSheetView `@outlineSymbols`) */
  outlineSymbols?: boolean;
  /** Show zero values (CT_CustomSheetView `@zeroValues`) */
  zeroValues?: boolean;
  /** Fit to page (CT_CustomSheetView `@fitToPage`) */
  fitToPage?: boolean;
  /** Print area (CT_CustomSheetView `@printArea`) */
  printArea?: boolean;
  /** Filter applied (CT_CustomSheetView `@filter`) */
  filter?: boolean;
  /** Show auto filter (CT_CustomSheetView `@showAutoFilter`) */
  showAutoFilter?: boolean;
  /** Hidden rows (CT_CustomSheetView `@hiddenRows`) */
  hiddenRows?: boolean;
  /** Hidden columns (CT_CustomSheetView `@hiddenColumns`) */
  hiddenColumns?: boolean;
  /** Sheet state (CT_CustomSheetView `@state`) */
  state?: "visible" | "hidden" | "veryHidden";
  /** Filter unique (CT_CustomSheetView `@filterUnique`) */
  filterUnique?: boolean;
  /** View type (CT_CustomSheetView `@view`) */
  view?: "normal" | "pageBreakPreview" | "pageLayout";
}

/** Cell watch entry (CT_CellWatch) */
export interface CellWatchOptions {
  /** Cell reference, e.g. "A1" (CT_CellWatch `@r`) */
  reference: string;
}

/** Data consolidation (CT_DataConsolidate) */
export interface DataConsolidateOptions {
  /** Consolidation function (CT_DataConsolidate `@function`) */
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
  /** Use top row labels (CT_DataConsolidate `@topLabels`) */
  topLabels?: boolean;
  /** Use left column labels (CT_DataConsolidate `@leftLabels`) */
  leftLabels?: boolean;
  /** Use start labels (CT_DataConsolidate `@startLabels`) */
  startLabels?: boolean;
  /** Link to source data (CT_DataConsolidate `@link`) */
  link?: boolean;
  /** Source data references */
  refs?: string[];
}

/**
 * Drawing in header/footer (CT_DrawingHF).
 *
 * The 18 counters are XML attribute names verbatim: first letter l/c/r is the
 * header/footer section (left/center/right), second h/f is header vs footer,
 * third o/e/f is the page flavor (odd/even/first). Each holds the 1-based
 * picture number within that slot (e.g. `lho` = 2nd picture in the left
 * section of the header on odd pages).
 */
export interface DrawingHfOptions {
  /**
   * Relationship ID for the drawing. Round-trip only: carried from the parsed
   * package's rels; a freshly generated workbook has no matching relationship
   * until an embedding API exists.
   */
  rId: string;
  /** Picture number, left header odd pages (`@lho`) */
  lho?: number;
  /** Picture number, left header even pages (`@lhe`) */
  lhe?: number;
  /** Picture number, left header first page (`@lhf`) */
  lhf?: number;
  /** Picture number, center header odd pages (`@cho`) */
  cho?: number;
  /** Picture number, center header even pages (`@che`) */
  che?: number;
  /** Picture number, center header first page (`@chf`) */
  chf?: number;
  /** Picture number, right header odd pages (`@rho`) */
  rho?: number;
  /** Picture number, right header even pages (`@rhe`) */
  rhe?: number;
  /** Picture number, right header first page (`@rhf`) */
  rhf?: number;
  /** Picture number, left footer odd pages (`@lfo`) */
  lfo?: number;
  /** Picture number, left footer even pages (`@lfe`) */
  lfe?: number;
  /** Picture number, left footer first page (`@lff`) */
  lff?: number;
  /** Picture number, center footer odd pages (`@cfo`) */
  cfo?: number;
  /** Picture number, center footer even pages (`@cfe`) */
  cfe?: number;
  /** Picture number, center footer first page (`@cff`) */
  cff?: number;
  /** Picture number, right footer odd pages (`@rfo`) */
  rfo?: number;
  /** Picture number, right footer even pages (`@rfe`) */
  rfe?: number;
  /** Picture number, right footer first page (`@rff`) */
  rff?: number;
}

export interface WorksheetOptions {
  name?: string;
  /** Workbook sheet id (CT_Sheet `@sheetId`) — unique but not necessarily sequential. */
  sheetId?: number;
  /** Visibility (CT_Sheet `@state`) */
  state?: "visible" | "hidden" | "veryHidden";
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
  /** Sheet-level sort state (CT_Worksheet `sortState` — sorting without a filter) */
  sortState?: SortStateOptions;
  images?: PictureOptions[];
  charts?: WorksheetChartOptions[];
  /** Anchored shapes (xdr:sp): geometry + optional text body. */
  shapes?: ShapeOptions[];
  /** Anchored connectors (xdr:cxnSp): line/arrow geometry. */
  connectors?: ConnectorOptions[];
  /** Anchored groups (xdr:grpSp): group transform + nested children. */
  groups?: GroupOptions[];
  dataValidations?: DataValidationOptions[];
  /** Disable data validation prompts (CT_DataValidations `@disablePrompts`) */
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
  phoneticPr?: PhoneticPropertiesOptions;
  /** Background image for the worksheet */
  backgroundImage?: SheetBackgroundImageOptions;
  /** Print options (CT_PrintOptions) */
  printOptions?: PrintOptions;
  /** Sheet format properties (CT_SheetFormatPr) */
  sheetFormatPr?: SheetFormatPropertiesOptions;
  /** Sheet extended properties (CT_SheetPr attributes) */
  sheetPr?: SheetPropertiesOptions;
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
  /** Drawing in header/footer (CT_DrawingHF) */
  drawingHF?: DrawingHfOptions;
  /**
   * Legacy drawing for header/footer r:id (CT_LegacyDrawingHF). Round-trip
   * only: the referenced VML part is not re-emitted by the compiler.
   */
  legacyDrawingHF?: string;
  /**
   * Drawing reference r:id (CT_Worksheet `<drawing>`). Round-trip only: a
   * drawing part whose anchors the bridge does not map onto options (e.g. OLE
   * object shape representations) passes through verbatim, so the original
   * reference stays valid. Absent on freshly authored worksheets — the
   * compiler derives the reference from images/charts/shapes.
   */
  drawingRid?: string;
  /**
   * Legacy drawing reference r:id (CT_Worksheet `<legacyDrawing>`). Round-trip
   * only, same passthrough semantics as {@link drawingRid}.
   */
  legacyDrawingRid?: string;
  /** Selections in sheet view (CT_Selection — one per pane, max 4) */
  selection?: SelectionOptions[];
  /** Pivot selection in sheet view (CT_PivotSelection) */
  pivotSelection?: PivotSelectionOptions;
  /** Cell smart tags (CT_SmartTags) */
  smartTags?: CellSmartTagsOptions[];
  /** Query tables on this sheet (xl/queryTables/queryTableN.xml) */
  queryTables?: QueryTableOptions[];
  /** Single-cell XML tables on this sheet (xl/tables/tableSingleCellsN.xml) */
  singleXmlCells?: SingleXmlCellOptions[];
  /** Sheet calc properties (CT_SheetCalcPr) */
  sheetCalcPr?: SheetCalculationPropertiesOptions;
  /** Extension list (extLst) */
  ext?: string;
  /** Control objects (CT_Controls) */
  controls?: ControlOptions[];
  /** Custom sheet properties (CT_CustomProperties) */
  customProperties?: CustomSheetPropertyOptions[];
  /** OLE objects (CT_OleObjects) */
  oleObjects?: OleObjectOptions[];
  /** Web publish items (CT_WebPublishItems) */
  webPublishItems?: WebPublishItemOptions[];
  /** Page margins in inches (CT_PageMargins) */
  pageMargins?: PageMarginsOptions;
  /** Cell range dimension (CT_Dimension ref); auto-computed when omitted */
  dimension?: string;
}

/** Page margins in inches (CT_PageMargins). Numbers are inches; strings are UniversalMeasure. */
export interface PageMarginsOptions {
  left?: number | UniversalMeasure;
  right?: number | UniversalMeasure;
  top?: number | UniversalMeasure;
  bottom?: number | UniversalMeasure;
  header?: number | UniversalMeasure;
  footer?: number | UniversalMeasure;
}

/** Sheet calc properties (CT_SheetCalcPr) */
export interface SheetCalculationPropertiesOptions {
  /** Full calc on load (CT_SheetCalcPr `@fullCalcOnLoad`) */
  fullCalcOnLoad?: boolean;
}

/** Form control object (CT_Control) */
export interface ControlOptions {
  /** Shape ID (CT_Control `@shapeId`) */
  shapeId: number;
  /**
   * Control r:id (CT_ControlPr `@r:id`). Round-trip only: the control's VML
   * and binary parts are not re-emitted, so the id is not resolvable in a
   * freshly generated workbook.
   */
  rId: string;
  /** Control name (CT_ControlPr `@name`) */
  name?: string;
  /** Locked (CT_ControlPr `@locked`) */
  locked?: boolean;
  /** UI-locked (CT_ControlPr `@uiObject`) */
  uiObject?: boolean;
  /** Recalc always (CT_ControlPr `@recalcAlways`) */
  recalcAlways?: boolean;
  /** Linked cell (CT_ControlPr `@linkedCell`) */
  linkedCell?: string;
  /** List fill range (CT_ControlPr `@listFillRange`) */
  listFillRange?: string;
  /** Control formula (CT_ControlPr `@cf`) */
  formula?: string;
  /** Use the default icon size (CT_ControlPr `@defaultSize`; default true). */
  defaultSize?: boolean;
  /** Auto line (CT_ControlPr `@autoLine`; default true). */
  autoLine?: boolean;
  /** Auto picture (CT_ControlPr `@autoPict`; default true). */
  autoPict?: boolean;
  /**
   * Relationship ID of the icon image (controlPr `@r:id`). Round-trip only:
   * the icon part is not re-emitted by the compiler.
   */
  iconRid?: string;
  /** Cell anchor inside controlPr (from/to corners, 0-based). */
  anchor?: ObjectAnchorOptions;
  /**
   * Source wrapped the control in mc:AlternateContent (Excel 2010+ form:
   * Choice carries the full element, Fallback the bare one). Re-emit the
   * wrapper only when the source had it.
   */
  alternateContent?: boolean;
}

/** Custom property (CT_CustomProperty) */
export interface CustomSheetPropertyOptions {
  /** Property name */
  name: string;
  /**
   * Relationship ID to the binary property part. Round-trip only: the part
   * is not re-emitted by the compiler.
   */
  rId: string;
}

/** OLE object (CT_OleObject) */
export interface OleObjectOptions {
  /** Program ID (CT_OleObject `@progId`) */
  progId?: string;
  /** Display aspect (CT_OleObject `@dvAspect`) */
  dvAspect?: "DVASPECT_CONTENT" | "DVASPECT_ICON";
  /** Linked source (CT_OleObject `@link`) */
  link?: string;
  /** OLE update mode (CT_OleObject `@oleUpdate`) */
  oleUpdate?: "OLEUPDATE_ALWAYS" | "OLEUPDATE_ONCALL";
  /** Auto load (CT_OleObject `@autoLoad`) */
  autoLoad?: boolean;
  /** Shape ID (CT_OleObject `@shapeId`) */
  shapeId: number;
  /**
   * Relationship ID (CT_OleObject `@r:id`). Round-trip only: the embedded
   * object part is not re-emitted, so the id is not resolvable in a freshly
   * generated workbook.
   */
  rId?: string;
  /** Object properties (CT_ObjectPr) */
  objectPr?: OleObjectPropertiesOptions;
  /**
   * Source wrapped the oleObject in mc:AlternateContent (Excel 2010+ form:
   * Choice carries the full element, Fallback the bare one). Re-emit the
   * wrapper only when the source had it.
   */
  alternateContent?: boolean;
}

/** OLE object properties (CT_ObjectPr) */
export interface OleObjectPropertiesOptions {
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
  /**
   * Relationship ID of the icon image (objectPr `@r:id`). Round-trip only:
   * the icon part is not re-emitted by the compiler.
   */
  iconRid?: string;
  /** Cell anchor inside objectPr (from/to corners, 0-based). */
  anchor?: ObjectAnchorOptions;
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
