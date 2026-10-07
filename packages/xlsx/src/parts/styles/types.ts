import type { ArgbHexColor, HexColor } from "@office-open/core";

import type { ColorOptions } from "../../shared/color";
/**
 * Styles — option types for xl/styles.xml.
 *
 * XLSX uses an index-based style system: cells reference style entries
 * via the `s` attribute, which is an index into `cellXfs`.
 *
 * @module
 */

// ── Sub-style option interfaces ──

export type FontUnderline = boolean | "single" | "double" | "singleAccounting" | "doubleAccounting";

export interface FontOptions {
  bold?: boolean;
  /** Source lexical form of `bold/@val`; round-trip only — do not hand-author. */
  boldRaw?: string;
  italic?: boolean;
  /** Source lexical form of `italic/@val`; round-trip only — do not hand-author. */
  italicRaw?: string;
  underline?: FontUnderline;
  strike?: boolean;
  /** Source lexical form of `strike/@val`; round-trip only — do not hand-author. */
  strikeRaw?: string;
  size?: number;
  /** Source lexical form of `size`; round-trip only — do not hand-author. */
  sizeRaw?: string;
  /** Font color (CT_Color) */
  color?: ColorOptions;
  font?: string;
  /** Character set (CT_Font/charset `@val`) */
  charset?: number;
  /** Font family (CT_Font/family `@val`) */
  family?: number;
  /** Condense (macOS, CT_Font/condense) */
  condense?: boolean;
  /** Source lexical form of `condense/@val`; round-trip only — do not hand-author. */
  condenseRaw?: string;
  /** Extend (macOS, CT_Font/extend) */
  extend?: boolean;
  /** Source lexical form of `extend/@val`; round-trip only — do not hand-author. */
  extendRaw?: string;
  /** Vertical alignment: superscript/subscript (CT_Font/vertAlign `@val`) */
  vertAlign?: "superscript" | "subscript" | "baseline";
  /** Font scheme (CT_Font/scheme `@val`) */
  scheme?: "major" | "minor" | "none";
  /** Font shadow (CT_Font/shadow) */
  shadow?: boolean;
  /** Source lexical form of `shadow/@val`; round-trip only — do not hand-author. */
  shadowRaw?: string;
  /** Font outline (CT_Font/outline) */
  outline?: boolean;
  /** Source lexical form of `outline/@val`; round-trip only — do not hand-author. */
  outlineRaw?: string;
  /** Source child element order (CT_Font children are an unbounded choice).
   * Round-trip only: writers emit children in this order when present. */
  childOrder?: string[];
}

/** Gradient stop (CT_GradientStop) */
export interface CellGradientStopOptions {
  /** Position (0.0–1.0) */
  position: number;
  /** Stop color (CT_Color) */
  color?: ColorOptions;
}

/**
 * Worksheet fill description. `color` alone is a solid fill; `patternType` +
 * `bgColor` describe a pattern fill; `stops` + `gradientType` describe a
 * gradient fill.
 */
export interface CellFillOptions {
  /** Fill kind; pattern/gradient read the dedicated fields below. */
  type?: "solid" | "pattern" | "gradient";
  /** Foreground color (CT_PatternFill/fgColor) */
  foregroundColor?: ColorOptions;
  /** Pattern kind (CT_PatternFill `@patternType`, ST_PatternType) */
  patternType?:
    | "none"
    | "solid"
    | "mediumGray"
    | "darkGray"
    | "lightGray"
    | "darkHorizontal"
    | "darkVertical"
    | "darkDown"
    | "darkUp"
    | "darkGrid"
    | "darkTrellis"
    | "lightHorizontal"
    | "lightVertical"
    | "lightDown"
    | "lightUp"
    | "lightGrid"
    | "lightTrellis"
    | "gray125"
    | "gray0625";
  /** Legacy SpreadsheetML `<pattern>` element spelling; round-trip only. */
  legacyPatternElement?: boolean;
  /** Source pattern kind when it is outside the modern ST_PatternType union. */
  legacyPatternType?: string;
  /** True when the source declared `@patternType`. Round-trip only. */
  patternTypeDeclared?: boolean;
  /**
   * Background color (CT_PatternFill/bgColor). Differential-format fills show
   * bgColor as the visible tint, so author `backgroundColor` for dxf fills.
   */
  backgroundColor?: ColorOptions;
  /** Gradient stops (CT_GradientFill/stop) */
  stops?: CellGradientStopOptions[];
  /** Gradient type (CT_GradientFill `@type`) */
  gradientType?: "linear" | "path";
  /** Gradient degree for linear (CT_GradientFill `@degree`) */
  gradientDegree?: number;
  /** Gradient left position for path (CT_GradientFill `@left`) */
  gradientLeft?: number;
  /** Gradient right position for path (CT_GradientFill `@right`) */
  gradientRight?: number;
  /** Gradient top position for path (CT_GradientFill `@top`) */
  gradientTop?: number;
  /** Gradient bottom position for path (CT_GradientFill `@bottom`) */
  gradientBottom?: number;
}

export interface BorderOptions {
  /** Border pattern (ST_BorderStyle); "medium*" variants thicken the pattern, "hair" is the thinnest hairline, "slantDashDot" slanted dashes with dots. */
  style?:
    | "none"
    | "thin"
    | "medium"
    | "dashed"
    | "dotted"
    | "thick"
    | "double"
    | "hair"
    | "mediumDashed"
    | "dashDot"
    | "mediumDashDot"
    | "dashDotDot"
    | "mediumDashDotDot"
    | "slantDashDot";
  /** Border color (CT_Color) */
  color?: ColorOptions;
}

/** The four sides plus diagonals of a cell border (CT_Border children). */
export interface BorderSideOptions {
  /** Source child element order (CT_Border children are a sequence with optional
   * RTL extensions). Round-trip only: writers emit children in this order. */
  childOrder?: string[];
  top?: BorderOptions;
  bottom?: BorderOptions;
  left?: BorderOptions;
  right?: BorderOptions;
  diagonal?: BorderOptions;
  /** Diagonal up (CT_Border `@diagonalUp`) — on the parent border element */
  diagonalUp?: boolean;
  /** Diagonal down (CT_Border `@diagonalDown`) — on the parent border element */
  diagonalDown?: boolean;
  /** Apply only the outside edges of a cell range (CT_Border `@outline`). */
  outline?: boolean;
  /** Leading edge border (CT_Border/start, for RTL support) */
  start?: BorderOptions;
  /** Trailing edge border (CT_Border/end, for RTL support) */
  end?: BorderOptions;
  /** Vertical inner border (CT_Border/vertical, for cell range borders) */
  vertical?: BorderOptions;
  /** Horizontal inner border (CT_Border/horizontal, for cell range borders) */
  horizontal?: BorderOptions;
}

export interface AlignmentOptions {
  /** Horizontal alignment: "general" per content type, "fill" repeat the value across the cell, "centerContinuous" merge with the cell to the right, "justify"/"distributed" spread words/columns evenly. */
  horizontal?:
    | "general"
    | "left"
    | "center"
    | "right"
    | "fill"
    | "justify"
    | "centerContinuous"
    | "distributed";
  /** Vertical alignment; "justify"/"distributed" spread lines evenly across the cell height. */
  vertical?: "top" | "center" | "bottom" | "justify" | "distributed";
  wrapText?: boolean;
  textRotation?: number;
  indent?: number;
  /** Source lexical form of `indent` (invalid producer values included);
   * round-trip only — do not hand-author. */
  indentRaw?: string;
  /** Relative indent (CT_CellAlignment `@relativeIndent`) */
  relativeIndent?: number;
  /** Justify last line (CT_CellAlignment `@justifyLastLine`) */
  justifyLastLine?: boolean;
  /** Shrink to fit (CT_CellAlignment `@shrinkToFit`) */
  shrinkToFit?: boolean;
  /** Reading order (CT_CellAlignment `@readingOrder`) */
  readingOrder?: number;
  /** Merge cells applies this alignment across the merged range (CT_CellAlignment `@mergeCell`). */
  mergeCell?: boolean;
}

export interface StyleOptions {
  font?: FontOptions;
  fill?: CellFillOptions;
  border?: BorderSideOptions;
  numFmt?: string;
  alignment?: AlignmentOptions;
  /** Quote prefix (CT_Xf `@quotePrefix`) */
  quotePrefix?: boolean;
  /** Pivot button (CT_Xf `@pivotButton`) */
  pivotButton?: boolean;
  /** Apply protection (CT_Xf `@applyProtection`) */
  applyProtection?: boolean;
  /** Cell protection (CT_CellProtection) */
  protection?: CellProtectionOptions;
}

/** Cell-level protection settings (CT_CellProtection) */
export interface CellProtectionOptions {
  /** Cell is locked (CT_CellProtection `@locked`) */
  locked?: boolean;
  /** Cell formula is hidden (CT_CellProtection `@hidden`) */
  hidden?: boolean;
}

/** Indexed color entry (CT_RgbColor) */
export interface IndexedColorOptions {
  /** RGB hex value, e.g. "FF000000" */
  rgb: ArgbHexColor;
}

/** Colors palette (CT_Colors) */
export interface ColorsOptions {
  /** Indexed color palette (CT_IndexedColors) */
  indexedColors?: IndexedColorOptions[];
  /** Legacy theme-palette integer channels; round-trip only. */
  themeColors?: number[];
  /** Most recently used colors (CT_MRUColors) */
  mruColors?: HexColor[];
  /** Legacy MRU channels (`type` + decimal `val`); round-trip only. */
  legacyMruColors?: { type: "rgb" | "theme" | "icv"; value: string }[];
}

/**
 * Inline number format inside a dxf (CT_NumFmt). `numFmtId` is required by the
 * XSD; when omitted the writer resolves it from the built-in table or falls
 * back to the custom range.
 */
export interface DxfNumFmtOptions {
  numFmtId?: number;
  formatCode: string;
}

/**
 * Differential format — conditional formatting changes, e.g.
 * `{ font: { color: "9C0006", bold: true }, fill: { color: "C6EFCE" } }`.
 */
export interface DxfOptions {
  font?: FontOptions;
  fill?: CellFillOptions;
  border?: BorderSideOptions;
  /** A plain string is shorthand for `{ formatCode: string }`. */
  numFmt?: string | DxfNumFmtOptions;
  alignment?: AlignmentOptions;
  protection?: CellProtectionOptions;
  /** Trailing dxf extensions (CT_Dxf/extLst) — round-trip only. */
  extensions?: StyleExtensionOptions[];
}

// ── Table / cell-style types ──

/** Table style element type (ST_TableStyleType). */
export type TableStyleElementType =
  | "wholeTable"
  | "headerRow"
  | "totalRow"
  | "firstColumn"
  | "lastColumn"
  | "firstRowStripe"
  | "secondRowStripe"
  | "firstColumnStripe"
  | "secondColumnStripe"
  | "firstHeaderCell"
  | "lastHeaderCell"
  | "firstTotalCell"
  | "lastTotalCell"
  | "subtotalRow1"
  | "subtotalRow2"
  | "subtotalRow3"
  | "subtotalColumn1"
  | "subtotalColumn2"
  | "subtotalColumn3"
  | "blankRow"
  | "firstColumnSubheading"
  | "secondColumnSubheading"
  | "thirdColumnSubheading"
  | "firstRowSubheading"
  | "secondRowSubheading"
  | "thirdRowSubheading"
  | "pageFieldLabels"
  | "pageFieldValues";

/** Table style element (CT_TableStyleElement). */
export interface TableStyleElementOptions {
  /** Element type */
  type: TableStyleElementType;
  /** Differential format index (dxf) */
  dxfId?: number;
  /** Stripe/band size (CT_TableStyleElement `@size`, default 1) */
  size?: number;
}

/** Custom table/pivot table style (CT_TableStyle). */
/** Style sheet extension (CT_Extension) */
export interface StyleExtensionOptions {
  /** Extension URI (required) */
  uri: string;
  /**
   * Namespace declarations carried on the ext element (xmlns:x14="…") — the
   * prefixed children are unbound without them.
   */
  namespaces?: Record<string, string>;
  /** Extension content (raw XML fragment) */
  content?: string;
}

/** A custom table or pivot style (styles.xml dxf-based tableStyle entry). */
export interface CustomTableStyleOptions {
  /** Style name (must be unique) */
  name: string;
  /** Revision uid (CT_TableStyle `@xr9:uid` etc.) */
  uid?: string;
  /** Pivot style (vs table style) */
  pivot?: boolean;
  /** Applies to tables (CT_TableStyle `@table`, default true) */
  table?: boolean;
  /** Table style elements */
  elements?: TableStyleElementOptions[];
}

/** Custom cell style (CT_CellStyle) — a named reference to a cell-style XF. */
export interface CustomCellStyleOptions {
  /** Style name */
  name: string;
  /** Index into cellStyleXfs; the referenced entry holds the format. */
  xfId: number;
  builtinId?: number;
  /** Custom built-in (CT_CellStyle `@customBuiltin`) */
  customBuiltin?: boolean;
  /** Outline level (CT_CellStyle `@iLevel`) */
  iLevel?: number;
  /** Hidden style (CT_CellStyle `@hidden`) */
  hidden?: boolean;
  /** Revision uid (CT_CellStyle `@xr:uid`) */
  uid?: string;
}

/** Cell XF entry exposed by Styles.toDescriptorOptions(). */
export interface CellXfEntry {
  fontId: number;
  fillId: number;
  borderId: number;
  numFmtId: number;
  alignment?: AlignmentOptions;
  quotePrefix?: boolean;
  pivotButton?: boolean;
  applyProtection?: boolean;
  protection?: CellProtectionOptions;
}

/**
 * Named cell-style template — a cellStyleXfs entry. Structured definitions
 * (not indices); source order keeps cellStyle.xfId stable on round-trip.
 * Alignment/protection and applyXxx flags preserved verbatim, unlike cellXfs
 * which derives applyXxx.
 */
export interface CellStyleXfOptions {
  font?: FontOptions;
  /** Source fontId when the index has no resolvable font (dangling reference);
   * round-trip only — do not hand-author. */
  fontIdRaw?: number;
  fill?: CellFillOptions;
  border?: BorderSideOptions;
  numFmt?: string;
  /** Raw numFmtId when the source id has no resolvable format code (round-trip only). */
  numFmtId?: number;
  /** True when the source xf declared `@fillId`. Round-trip only: attribute
   * presence and absence stay distinct in rebuilt styles.xml. */
  fillIdDeclared?: boolean;
  /** True when the source xf declared `@borderId`. Round-trip only. */
  borderIdDeclared?: boolean;
  alignment?: AlignmentOptions;
  protection?: CellProtectionOptions;
  quotePrefix?: boolean;
  pivotButton?: boolean;
  applyNumberFormat?: boolean;
  applyFont?: boolean;
  applyFill?: boolean;
  applyBorder?: boolean;
  applyAlignment?: boolean;
  applyProtection?: boolean;
}

/** Fonts container metadata (CT_Fonts `@x14ac:knownFonts`). */
export interface FontsContainerOptions {
  /** Whether Excel marked the font table as known (CT_Fonts `@x14ac:knownFonts`). */
  knownFonts?: boolean;
}

/** Snapshot of Styles internal state for descriptor-based XML generation. */
export interface StylesState {
  customNumFmts: readonly NumFmtEntry[];
  fonts: FontOptions[];
  fills: CellFillOptions[];
  borders: BorderSideOptions[];
  cellXfs: CellXfEntry[];
  dxfs: DxfOptions[];
  colors?: ColorsOptions;
  tableStyles?: CustomTableStyleOptions[];
  customCellStyles?: CustomCellStyleOptions[];
  styleExtensions?: StyleExtensionOptions[];
}

/**
 * Indexed XF reference produced by `stylesDesc`.parse — index-based
 * (fontId/fillId/…) rather than resolved objects, consumed by callers that
 * resolve indices into fonts/fills/borders arrays.
 */
export interface IndexedXfEntry {
  fontId?: number;
  fillId?: number;
  borderId?: number;
  numFmtId?: number;
  /** True when the source declared `@fontId`. Round-trip only. */
  fontIdDeclared?: boolean;
  /** True when the source declared `@fillId`. Round-trip only. */
  fillIdDeclared?: boolean;
  /** True when the source declared `@borderId`. Round-trip only. */
  borderIdDeclared?: boolean;
  /** True when the source declared `@numFmtId`. Round-trip only. */
  numFmtIdDeclared?: boolean;
  /** Index into cellStyleXfs this xf derives from (CT_Xf/@xfId) */
  xfId?: number;
  /** True when the source declared `@xfId`. Round-trip only. */
  xfIdDeclared?: boolean;
  alignment?: AlignmentOptions;
  protection?: CellProtectionOptions;
  quotePrefix?: boolean;
  pivotButton?: boolean;
  /**
   * Explicit apply* flags from the source xf, preserved verbatim on
   * round-trip. Undefined means the source omitted them — stringify derives
   * them instead (the fresh-generation behavior).
   */
  applyFont?: boolean;
  applyFill?: boolean;
  applyBorder?: boolean;
  applyNumberFormat?: boolean;
  applyAlignment?: boolean;
  applyProtection?: boolean;
}

/** Table styles block (CT_TableStyles), including its container attributes. */
export interface TableStylesInfo {
  /** Declared number of custom styles (CT_TableStyles `@count`) */
  count?: number;
  /** Default style applied to new tables (CT_TableStyles `@defaultTableStyle`) */
  defaultTableStyle?: string;
  /** Default style applied to new PivotTables (CT_TableStyles `@defaultPivotStyle`) */
  defaultPivotStyle?: string;
  /** Custom table and PivotTable styles in document order */
  tableStyles?: CustomTableStyleOptions[];
}

/** A numFmts section entry (CT_NumFmt), as written in the source. */
export interface NumFmtEntry {
  numFmtId: number;
  formatCode: string;
}

/** Result of {@link stylesDesc}.parse (xl/styles.xml → structured data). */
export interface StylesParseResult {
  /** numFmts section entries in document order (for table adoption). */
  numFmts?: NumFmtEntry[];
  fonts?: FontOptions[];
  fontsContainer?: FontsContainerOptions;
  fills?: CellFillOptions[];
  borders?: BorderSideOptions[];
  cellStyleXfs?: CellStyleXfOptions[];
  cellXfs?: IndexedXfEntry[];
  customCellStyles?: CustomCellStyleOptions[];
  dxfs?: DxfOptions[];
  tableStylesInfo?: TableStylesInfo;
  colors?: ColorsOptions;
  styleExtensions?: StyleExtensionOptions[];
}

// ── Descriptor Types ──

export interface StylesDocOptions {
  /** The Styles accumulator instance (for stringify). */
  styles: import("./styles").Styles;
}
