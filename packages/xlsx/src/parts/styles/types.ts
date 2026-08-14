/**
 * Styles — option types for xl/styles.xml.
 *
 * XLSX uses an index-based style system: cells reference style entries
 * via the `s` attribute, which is an index into `cellXfs`.
 *
 * @module
 */

// ── Sub-style option interfaces ──

export interface FontOptions {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  size?: number;
  color?: string;
  font?: string;
  /** Character set (CT_Font/charset @val) */
  charset?: number;
  /** Font family (CT_Font/family @val) */
  family?: number;
  /** Condense (macOS, CT_Font/condense) */
  condense?: boolean;
  /** Extend (macOS, CT_Font/extend) */
  extend?: boolean;
  /** Vertical alignment: superscript/subscript (CT_Font/vertAlign @val) */
  vertAlign?: "superscript" | "subscript" | "baseline";
  /** Font scheme (CT_Font/scheme @val) */
  scheme?: "major" | "minor" | "none";
  /** Font shadow (CT_Font/shadow) */
  shadow?: boolean;
  /** Font outline (CT_Font/outline) */
  outline?: boolean;
}

/** Gradient stop (CT_GradientStop) */
export interface CellGradientStopOptions {
  /** Position (0.0–1.0) */
  position: number;
  /** RGB color hex without alpha, e.g. "FF0000" */
  color: string;
}

export interface CellFillOptions {
  type?: "solid" | "pattern" | "gradient";
  color?: string;
  patternType?: string;
  /** Background color for pattern fill (CT_PatternFill/bgColor) */
  bgColor?: string;
  /** Background color indexed (CT_Color @indexed) */
  colorIndexed?: number;
  /** Gradient stops (CT_GradientFill/stop) */
  stops?: CellGradientStopOptions[];
  /** Gradient type (CT_GradientFill @type) */
  gradientType?: "linear" | "path";
  /** Gradient degree for linear (CT_GradientFill @degree) */
  gradientDegree?: number;
  /** Gradient left position for path (CT_GradientFill @left) */
  gradientLeft?: number;
  /** Gradient right position for path (CT_GradientFill @right) */
  gradientRight?: number;
  /** Gradient top position for path (CT_GradientFill @top) */
  gradientTop?: number;
  /** Gradient bottom position for path (CT_GradientFill @bottom) */
  gradientBottom?: number;
}

export interface BorderOptions {
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
  color?: string;
}

export interface BorderSideOptions {
  top?: BorderOptions;
  bottom?: BorderOptions;
  left?: BorderOptions;
  right?: BorderOptions;
  diagonal?: BorderOptions;
  /** Diagonal up (CT_Border @diagonalUp) — on the parent border element */
  diagonalUp?: boolean;
  /** Diagonal down (CT_Border @diagonalDown) — on the parent border element */
  diagonalDown?: boolean;
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
  horizontal?:
    | "general"
    | "left"
    | "center"
    | "right"
    | "fill"
    | "justify"
    | "centerContinuous"
    | "distributed";
  vertical?: "top" | "center" | "bottom" | "justify" | "distributed";
  wrapText?: boolean;
  textRotation?: number;
  indent?: number;
  /** Relative indent (CT_CellAlignment @relativeIndent) */
  relativeIndent?: number;
  /** Justify last line (CT_CellAlignment @justifyLastLine) */
  justifyLastLine?: boolean;
  /** Shrink to fit (CT_CellAlignment @shrinkToFit) */
  shrinkToFit?: boolean;
  /** Reading order (CT_CellAlignment @readingOrder) */
  readingOrder?: number;
}

export interface StyleOptions {
  font?: FontOptions;
  fill?: CellFillOptions;
  border?: BorderSideOptions;
  numFmt?: string;
  alignment?: AlignmentOptions;
  /** Quote prefix (CT_Xf @quotePrefix) */
  quotePrefix?: boolean;
  /** Pivot button (CT_Xf @pivotButton) */
  pivotButton?: boolean;
  /** Apply protection (CT_Xf @applyProtection) */
  applyProtection?: boolean;
  /** Cell protection (CT_CellProtection) */
  protection?: CellProtectionOptions;
}

/** Cell-level protection settings (CT_CellProtection) */
export interface CellProtectionOptions {
  /** Cell is locked (CT_CellProtection @locked) */
  locked?: boolean;
  /** Cell formula is hidden (CT_CellProtection @hidden) */
  hidden?: boolean;
}

/** Indexed color entry (CT_RgbColor) */
export interface IndexedColorOptions {
  /** RGB hex value, e.g. "FF000000" */
  rgb: string;
}

/** Colors palette (CT_Colors) */
export interface ColorsOptions {
  /** Indexed color palette (CT_IndexedColors) */
  indexedColors?: IndexedColorOptions[];
  /** Most recently used colors (CT_MRUColors) */
  mruColors?: string[];
}

/** Differential format — used by conditional formatting to specify what changes. */
export interface DxfOptions {
  font?: FontOptions;
  fill?: CellFillOptions;
  border?: BorderSideOptions;
  numFmt?: string;
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
  /** Button style (for pivot tables) */
  button?: boolean;
}

/** Custom table/pivot table style (CT_TableStyle). */
/** Style sheet extension (CT_Extension) */
export interface StyleExtensionOptions {
  /** Extension URI (required) */
  uri: string;
  /** Extension content (raw XML fragment) */
  content?: string;
}

export interface CustomTableStyleOptions {
  /** Style name (must be unique) */
  name: string;
  /** Pivot style (vs table style) */
  pivot?: boolean;
  /** Table style elements */
  elements?: TableStyleElementOptions[];
}

/** Custom cell style (CT_CellStyle) */
export interface CustomCellStyleOptions {
  /** Style name */
  name: string;
  /** XF index to apply */
  xfId: number;
  /** Built-in ID */
  builtinId?: number;
  /** Custom built-in (CT_CellStyle @customBuiltin) */
  customBuiltin?: boolean;
  /** Outline level (CT_CellStyle @iLevel) */
  iLevel?: number;
  /** Hidden style (CT_CellStyle @hidden) */
  hidden?: boolean;
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
 * Named cell-style template — a cellStyleXfs entry (CT_Xf in the cellStyleXfs
 * context). Structured (font/fill/border/numFmt definitions, not indices) so
 * the compiler can re-register each definition against its rebuilt tables when
 * round-tripping; cellStyle.xfId references stay stable because entries keep
 * their source order. Alignment/protection and the applyXxx flags are preserved
 * verbatim (named-style fidelity), unlike cellXfs which derives applyXxx.
 */
export interface CellStyleXfOptions {
  font?: FontOptions;
  fill?: CellFillOptions;
  border?: BorderSideOptions;
  numFmt?: string;
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

/** Snapshot of Styles internal state for descriptor-based XML generation. */
export interface StylesState {
  customNumFmts: ReadonlyMap<string, number>;
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
 * Indexed XF reference produced by {@link stylesDesc}.parse — index-based
 * (fontId/fillId/…) rather than resolved objects, consumed by callers that
 * resolve indices into fonts/fills/borders arrays.
 */
export interface IndexedXfEntry {
  fontId?: number;
  fillId?: number;
  borderId?: number;
  numFmtId?: number;
  alignment?: AlignmentOptions;
  protection?: CellProtectionOptions;
  quotePrefix?: boolean;
  pivotButton?: boolean;
}

/** Table styles block (CT_TableStyles) produced by {@link stylesDesc}.parse. */
export interface TableStylesInfo {
  count?: number;
  defaultTableStyle?: string;
  defaultPivotStyle?: string;
  tableStyles?: CustomTableStyleOptions[];
}

/** Result of {@link stylesDesc}.parse (xl/styles.xml → structured data). */
export interface StylesParseResult {
  /** Reverse map numFmtId → formatCode, for O(1) lookup in resolveStyle. */
  customNumFmtById?: Map<number, string>;
  fonts?: FontOptions[];
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
