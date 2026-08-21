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
  /**
   * Theme palette index (CT_Color `@theme`) — takes precedence over `color`
   * when both are set, matching the XSD's single-channel choice.
   */
  themeColor?: number;
  /** Tint applied to the theme color (CT_Color `@tint`) */
  tint?: number;
  /** Indexed color palette entry (CT_Color `@indexed`) */
  colorIndexed?: number;
  /** Automatic (system) color instead of an explicit RGB (CT_Color `@auto`) */
  autoColor?: boolean;
  font?: string;
  /** Character set (CT_Font/charset `@val`) */
  charset?: number;
  /** Font family (CT_Font/family `@val`) */
  family?: number;
  /** Condense (macOS, CT_Font/condense) */
  condense?: boolean;
  /** Extend (macOS, CT_Font/extend) */
  extend?: boolean;
  /** Vertical alignment: superscript/subscript (CT_Font/vertAlign `@val`) */
  vertAlign?: "superscript" | "subscript" | "baseline";
  /** Font scheme (CT_Font/scheme `@val`) */
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

/**
 * Worksheet fill description. `color` alone is a solid fill; `patternType` +
 * `bgColor` describe a pattern fill; `stops` + `gradientType` describe a
 * gradient fill.
 */
export interface CellFillOptions {
  type?: "solid" | "pattern" | "gradient";
  /** Foreground color hex without alpha, e.g. "C6EFCE" */
  color?: string;
  /** Foreground theme palette index (CT_Color `@theme` on fgColor) */
  themeColor?: number;
  /** Foreground tint (CT_Color `@tint` on fgColor) */
  tint?: number;
  /** Foreground automatic color (CT_Color `@auto` on fgColor) */
  fgAutoColor?: boolean;
  patternType?: string;
  /** Background color for pattern fill (CT_PatternFill/bgColor) */
  bgColor?: string;
  /** Background theme palette index (CT_Color `@theme` on bgColor) */
  bgThemeColor?: number;
  /** Background tint (CT_Color `@tint` on bgColor) */
  bgTint?: number;
  /** Foreground color indexed (CT_Color `@indexed` on fgColor) */
  colorIndexed?: number;
  /** Background color indexed (CT_Color `@indexed` on bgColor) */
  bgColorIndexed?: number;
  /** Background automatic color (CT_Color `@auto` on bgColor) */
  bgAutoColor?: boolean;
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
  /** Theme palette index (CT_Color `@theme`) — takes precedence over `color` */
  themeColor?: number;
  /** Tint applied to the theme color (CT_Color `@tint`) */
  tint?: number;
  /** Automatic (system) color instead of an explicit RGB (CT_Color `@auto`) */
  autoColor?: boolean;
  /** Indexed color palette entry (CT_Color `@indexed`) */
  colorIndexed?: number;
}

export interface BorderSideOptions {
  top?: BorderOptions;
  bottom?: BorderOptions;
  left?: BorderOptions;
  right?: BorderOptions;
  diagonal?: BorderOptions;
  /** Diagonal up (CT_Border `@diagonalUp`) — on the parent border element */
  diagonalUp?: boolean;
  /** Diagonal down (CT_Border `@diagonalDown`) — on the parent border element */
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
  /** Relative indent (CT_CellAlignment `@relativeIndent`) */
  relativeIndent?: number;
  /** Justify last line (CT_CellAlignment `@justifyLastLine`) */
  justifyLastLine?: boolean;
  /** Shrink to fit (CT_CellAlignment `@shrinkToFit`) */
  shrinkToFit?: boolean;
  /** Reading order (CT_CellAlignment `@readingOrder`) */
  readingOrder?: number;
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
  rgb: string;
}

/** Colors palette (CT_Colors) */
export interface ColorsOptions {
  /** Indexed color palette (CT_IndexedColors) */
  indexedColors?: IndexedColorOptions[];
  /** Most recently used colors (CT_MRUColors) */
  mruColors?: string[];
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
 * Differential format — used by conditional formatting to specify what changes.
 * Example: `{ font: { color: "9C0006", bold: true }, fill: { color: "C6EFCE" } }`.
 */
export interface DxfOptions {
  font?: FontOptions;
  fill?: CellFillOptions;
  border?: BorderSideOptions;
  /** A plain string is shorthand for `{ formatCode: string }`. */
  numFmt?: string | DxfNumFmtOptions;
  alignment?: AlignmentOptions;
  protection?: CellProtectionOptions;
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
  /**
   * Namespace declarations carried on the ext element (xmlns:x14="…") — the
   * prefixed children are unbound without them.
   */
  namespaces?: Record<string, string>;
  /** Extension content (raw XML fragment) */
  content?: string;
}

export interface CustomTableStyleOptions {
  /** Style name (must be unique) */
  name: string;
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
  /** Built-in ID */
  builtinId?: number;
  /** Custom built-in (CT_CellStyle `@customBuiltin`) */
  customBuiltin?: boolean;
  /** Outline level (CT_CellStyle `@iLevel`) */
  iLevel?: number;
  /** Hidden style (CT_CellStyle `@hidden`) */
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
  /** Index into cellStyleXfs this xf derives from (CT_Xf/@xfId) */
  xfId?: number;
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

/** Table styles block (CT_TableStyles) produced by {@link stylesDesc}.parse. */
export interface TableStylesInfo {
  count?: number;
  defaultTableStyle?: string;
  defaultPivotStyle?: string;
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
