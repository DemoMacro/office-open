/**
 * Styles component — generates xl/styles.xml.
 *
 * XLSX uses an index-based style system: cells reference style entries
 * via the `s` attribute, which is an index into `cellXfs`.
 *
 * @module
 */
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attrs, escapeXml, findChild, attr, attrNum } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

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
export interface GradientStopOptions {
  /** Position (0.0–1.0) */
  position: number;
  /** RGB color hex without alpha, e.g. "FF0000" */
  color: string;
}

export interface FillOptions {
  type?: "solid" | "pattern" | "gradient";
  color?: string;
  patternType?: string;
  /** Background color for pattern fill (CT_PatternFill/bgColor) */
  bgColor?: string;
  /** Background color indexed (CT_Color @indexed) */
  colorIndexed?: number;
  /** Gradient stops (CT_GradientFill/stop) */
  stops?: GradientStopOptions[];
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
  style?: "thin" | "medium" | "thick" | "dotted" | "dashed" | "hair" | "none";
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
  horizontal?: "left" | "center" | "right" | "fill" | "justify";
  vertical?: "top" | "center" | "bottom";
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
  fill?: FillOptions;
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
  fill?: FillOptions;
  border?: BorderSideOptions;
  numFmt?: string;
}

// ── Style key helpers for deduplication ──

function fontKey(f: FontOptions): string {
  return `b${f.bold ? 1 : 0}i${f.italic ? 1 : 0}u${f.underline ? 1 : 0}s${f.strike ? 1 : 0}z${f.size ?? 0}c${f.color ?? ""}n${f.font ?? ""}cs${f.charset ?? ""}fm${f.family ?? ""}co${f.condense ? 1 : 0}ex${f.extend ? 1 : 0}va${f.vertAlign ?? ""}sc${f.scheme ?? ""}sh${f.shadow ? 1 : 0}ol${f.outline ? 1 : 0}`;
}

function fillKey(f: FillOptions): string {
  return `t${f.type ?? ""}c${f.color ?? ""}p${f.patternType ?? ""}bg${f.bgColor ?? ""}g${f.stops?.map((s) => `${s.position}_${s.color}`).join("|") ?? ""}`;
}

function borderKey(b: BorderSideOptions): string {
  const sk = (o?: BorderOptions) => `${o?.style ?? ""}_${o?.color ?? ""}`;
  return `t${sk(b.top)}b${sk(b.bottom)}l${sk(b.left)}r${sk(b.right)}d${sk(b.diagonal)}du${b.diagonalUp ? 1 : 0}dd${b.diagonalDown ? 1 : 0}st${sk(b.start)}en${sk(b.end)}v${sk(b.vertical)}h${sk(b.horizontal)}`;
}

// ── Built-in number format IDs ──

const BUILTIN_NUMFMTS: Record<string, number> = {
  General: 0,
  "0": 1,
  "0.00": 2,
  "#,##0": 3,
  "#,##0.00": 4,
  "0%": 9,
  "0.00%": 10,
  "0.00E+00": 11,
  "mm-dd-yy": 14,
  "d-mmm-yy": 15,
  "d-mmm": 16,
  "mmm-yy": 17,
  "h:mm AM/PM": 18,
  "h:mm:ss AM/PM": 19,
  "h:mm": 20,
  "h:mm:ss": 21,
  "m/d/yy h:mm": 22,
  "#,##0 ;(#,##0)": 37,
  "#,##0 ;[Red](#,##0)": 38,
  "#,##0.00;(#,##0.00)": 39,
  "#,##0.00;[Red](#,##0.00)": 40,
  "mm:ss": 45,
  "[h]:mm:ss": 46,
  "mmss.0": 47,
  "##0.0E+0": 48,
  "@": 49,
};

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

/** Snapshot of Styles internal state for descriptor-based XML generation. */
export interface StylesState {
  customNumFmts: ReadonlyMap<string, number>;
  fonts: FontOptions[];
  fills: FillOptions[];
  borders: BorderSideOptions[];
  cellXfs: CellXfEntry[];
  dxfs: DxfOptions[];
  colors?: ColorsOptions;
  tableStyles?: CustomTableStyleOptions[];
  customCellStyles?: CustomCellStyleOptions[];
  styleExtensions?: StyleExtensionOptions[];
}

export class Styles {
  private fonts: FontOptions[] = [
    { size: 11, font: "Calibri" }, // default font (index 0)
  ];
  private fontKeys = new Map<string, number>();

  private fills: FillOptions[] = [
    { patternType: "none" }, // default fill (index 0)
    { patternType: "gray125" }, // required fill (index 1)
  ];
  private fillKeys = new Map<string, number>();

  private borders: BorderSideOptions[] = [
    {}, // default empty border (index 0)
  ];
  private borderKeys = new Map<string, number>();

  private customNumFmts = new Map<string, number>();
  private nextCustomNumFmtId = 164; // custom numFmts start at 164

  private cellXfs: Array<{
    fontId: number;
    fillId: number;
    borderId: number;
    numFmtId: number;
    alignment?: AlignmentOptions;
    quotePrefix?: boolean;
    pivotButton?: boolean;
    applyProtection?: boolean;
    protection?: CellProtectionOptions;
  }> = [
    { fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 }, // default xf (index 0)
  ];
  private cellXfKeys = new Map<string, number>();

  private dxfs: DxfOptions[] = [];

  private colors?: ColorsOptions;
  private tableStyles?: CustomTableStyleOptions[];
  /** Custom cell styles (CT_CellStyles) */
  private customCellStyles?: CustomCellStyleOptions[];
  /** Style sheet extensions (CT_ExtensionList) */
  private styleExtensions?: StyleExtensionOptions[];

  public constructor() {
    // Pre-register default font/fill/border keys
    this.fontKeys.set(fontKey(this.fonts[0]), 0);
    this.fillKeys.set(fillKey(this.fills[0]), 0);
    this.fillKeys.set(fillKey(this.fills[1]), 1);
    this.borderKeys.set(borderKey(this.borders[0]), 0);
    this.cellXfKeys.set(this.cellXfKey(this.cellXfs[0]), 0);
  }

  /**
   * Register a style and return its index (for the cell `s` attribute).
   * Deduplicates across fonts, fills, borders, numFmts, and cellXfs.
   */
  public register(opts: StyleOptions): number {
    const fontId = this.registerFont(opts.font);
    const fillId = this.registerFill(opts.fill);
    const borderId = this.registerBorder(opts.border);
    const numFmtId = this.registerNumFmt(opts.numFmt);

    const xf = {
      fontId,
      fillId,
      borderId,
      numFmtId,
      alignment: opts.alignment,
      quotePrefix: opts.quotePrefix,
      pivotButton: opts.pivotButton,
      applyProtection: opts.applyProtection,
      protection: opts.protection,
    };

    const key = this.cellXfKey(xf);
    const existing = this.cellXfKeys.get(key);
    if (existing !== undefined) return existing;

    const idx = this.cellXfs.length;
    this.cellXfs.push(xf);
    this.cellXfKeys.set(key, idx);
    return idx;
  }

  /**
   * Register a differential format and return its index (dxfId).
   * Used by conditional formatting rules.
   */
  public registerDxf(opts: DxfOptions): number {
    const idx = this.dxfs.length;
    this.dxfs.push(opts);
    return idx;
  }

  /**
   * Set color palette (indexed colors and MRU colors).
   */
  public setColors(opts: ColorsOptions): void {
    this.colors = opts;
  }

  public setTableStyles(styles: CustomTableStyleOptions[]): void {
    this.tableStyles = styles;
  }

  public setExtensions(extensions: StyleExtensionOptions[]): void {
    this.styleExtensions = extensions;
  }

  public setCustomCellStyles(styles: CustomCellStyleOptions[]): void {
    this.customCellStyles = styles;
  }

  /**
   * Expose internal state for descriptor-based XML generation.
   * The descriptor reads this snapshot to produce xl/styles.xml.
   */
  public toDescriptorOptions(): StylesState {
    return {
      customNumFmts: new Map(this.customNumFmts),
      fonts: [...this.fonts],
      fills: [...this.fills],
      borders: [...this.borders],
      cellXfs: [...this.cellXfs],
      dxfs: [...this.dxfs],
      colors: this.colors,
      tableStyles: this.tableStyles,
      customCellStyles: this.customCellStyles,
      styleExtensions: this.styleExtensions,
    };
  }

  private registerFont(opts?: FontOptions): number {
    if (!opts) return 0;
    const key = fontKey(opts);
    const existing = this.fontKeys.get(key);
    if (existing !== undefined) return existing;

    const idx = this.fonts.length;
    this.fonts.push(opts);
    this.fontKeys.set(key, idx);
    return idx;
  }

  private registerFill(opts?: FillOptions): number {
    if (!opts) return 0;
    const key = fillKey(opts);
    const existing = this.fillKeys.get(key);
    if (existing !== undefined) return existing;

    const idx = this.fills.length;
    this.fills.push(opts);
    this.fillKeys.set(key, idx);
    return idx;
  }

  private registerBorder(opts?: BorderSideOptions): number {
    if (!opts) return 0;
    const key = borderKey(opts);
    const existing = this.borderKeys.get(key);
    if (existing !== undefined) return existing;

    const idx = this.borders.length;
    this.borders.push(opts);
    this.borderKeys.set(key, idx);
    return idx;
  }

  private registerNumFmt(fmt?: string): number {
    if (!fmt) return 0;
    const builtin = BUILTIN_NUMFMTS[fmt];
    if (builtin !== undefined) return builtin;

    const existing = this.customNumFmts.get(fmt);
    if (existing !== undefined) return existing;

    const id = this.nextCustomNumFmtId++;
    this.customNumFmts.set(fmt, id);
    return id;
  }

  private cellXfKey(xf: {
    fontId: number;
    fillId: number;
    borderId: number;
    numFmtId: number;
    alignment?: AlignmentOptions;
    quotePrefix?: boolean;
    pivotButton?: boolean;
    applyProtection?: boolean;
    protection?: CellProtectionOptions;
  }): string {
    const a = xf.alignment;
    const ak = a
      ? `h${a.horizontal ?? ""}v${a.vertical ?? ""}w${a.wrapText ? 1 : 0}r${a.textRotation ?? ""}i${a.indent ?? ""}ri${a.relativeIndent ?? ""}jl${a.justifyLastLine ? 1 : 0}st${a.shrinkToFit ? 1 : 0}ro${a.readingOrder ?? ""}`
      : "";
    const pr = xf.protection;
    const pk = pr ? `l${pr.locked ?? ""}h${pr.hidden ?? ""}` : "";
    return `${xf.fontId}|${xf.fillId}|${xf.borderId}|${xf.numFmtId}|${ak}|qp${xf.quotePrefix ? 1 : 0}|pb${xf.pivotButton ? 1 : 0}|${pk}`;
  }

  // ── XML generation ──

  /**
   * Zero-allocation fast path: directly concatenate XML string.
   * Bypasses the IXmlableObject intermediate tree entirely.
   */
  /** Serialize to xl/styles.xml content (without XML declaration). */
  public serialize(): string {
    const p: string[] = [
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    ];

    // numFmts
    if (this.customNumFmts.size > 0) {
      p.push(`<numFmts count="${this.customNumFmts.size}">`);
      for (const [fmt, id] of this.customNumFmts) {
        p.push(`<numFmt numFmtId="${id}" formatCode="${escapeXml(fmt)}"/>`);
      }
      p.push("</numFmts>");
    }

    // fonts
    p.push(`<fonts count="${this.fonts.length}">`);
    for (const f of this.fonts) {
      p.push(`<font>${this.fontXmlStr(f)}</font>`);
    }
    p.push("</fonts>");

    // fills
    p.push(`<fills count="${this.fills.length}">`);
    for (const f of this.fills) {
      if (f.type === "gradient" && f.stops && f.stops.length > 0) {
        const gfAttrs: Record<string, string | number | boolean | undefined> = {};
        if (f.gradientType && f.gradientType !== "linear") gfAttrs.type = f.gradientType;
        if (f.gradientDegree !== undefined) gfAttrs.degree = f.gradientDegree;
        if (f.gradientLeft !== undefined) gfAttrs.left = f.gradientLeft;
        if (f.gradientRight !== undefined) gfAttrs.right = f.gradientRight;
        if (f.gradientTop !== undefined) gfAttrs.top = f.gradientTop;
        if (f.gradientBottom !== undefined) gfAttrs.bottom = f.gradientBottom;
        const stopParts = f.stops
          .map((s) => `<stop position="${s.position}"><color rgb="FF${s.color}"/></stop>`)
          .join("");
        p.push(`<fill><gradientFill${attrs(gfAttrs)}>${stopParts}</gradientFill></fill>`);
      } else {
        const patternAttrs = attrs({ patternType: f.patternType ?? "solid" });
        const fgColor = f.color
          ? `<fgColor rgb="FF${f.color}"/>`
          : f.colorIndexed !== undefined
            ? `<fgColor indexed="${f.colorIndexed}"/>`
            : "";
        const bgColor = f.bgColor ? `<bgColor rgb="FF${f.bgColor}"/>` : "";
        const colorContent = fgColor + bgColor;
        p.push(
          colorContent
            ? `<fill><patternFill${patternAttrs}>${colorContent}</patternFill></fill>`
            : `<fill><patternFill${patternAttrs}/></fill>`,
        );
      }
    }
    p.push("</fills>");

    // borders
    p.push(`<borders count="${this.borders.length}">`);
    for (const b of this.borders) {
      const bAttrs: string[] = [];
      if (b.diagonalUp) bAttrs.push('diagonalUp="1"');
      if (b.diagonalDown) bAttrs.push('diagonalDown="1"');
      const bAttr = bAttrs.length ? ` ${bAttrs.join(" ")}` : "";
      p.push(`<border${bAttr}>${this.borderXmlStr(b)}</border>`);
    }
    p.push("</borders>");

    // cellStyleXfs
    p.push(
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
    );

    // cellXfs
    p.push(`<cellXfs count="${this.cellXfs.length}">`);
    for (const xf of this.cellXfs) {
      const xAttrs: Record<string, string | number | boolean | undefined> = {
        numFmtId: xf.numFmtId,
        fontId: xf.fontId,
        fillId: xf.fillId,
        borderId: xf.borderId,
        xfId: 0,
      };
      if (xf.alignment) xAttrs.applyAlignment = 1;
      if (xf.fontId > 0) xAttrs.applyFont = 1;
      if (xf.fillId > 0) xAttrs.applyFill = 1;
      if (xf.borderId > 0) xAttrs.applyBorder = 1;
      if (xf.numFmtId > 0) xAttrs.applyNumberFormat = 1;
      if (xf.quotePrefix) xAttrs.quotePrefix = 1;
      if (xf.pivotButton) xAttrs.pivotButton = 1;
      if (xf.applyProtection) xAttrs.applyProtection = 1;
      if (xf.protection) xAttrs.applyProtection = xAttrs.applyProtection ?? 1;

      const alignStr = xf.alignment ? this.alignmentXmlStr(xf.alignment) : "";
      const protStr = xf.protection ? this.protectionXmlStr(xf.protection) : "";
      const inner = alignStr + protStr;
      p.push(inner ? `<xf${attrs(xAttrs)}>${inner}</xf>` : `<xf${attrs(xAttrs)}/>`);
    }
    p.push("</cellXfs>");

    // cellStyles
    if (this.customCellStyles && this.customCellStyles.length > 0) {
      const csAttrs: string[] = [`count="${this.customCellStyles.length + 1}"`];
      const csParts: string[] = [`<cellStyles ${csAttrs.join(" ")}>`];
      csParts.push('<cellStyle name="Normal" xfId="0" builtinId="0"/>');
      for (const cs of this.customCellStyles) {
        const attrs: string[] = [`name="${escapeXml(cs.name)}"`, `xfId="${cs.xfId}"`];
        if (cs.builtinId !== undefined) attrs.push(`builtinId="${cs.builtinId}"`);
        if (cs.customBuiltin) attrs.push('customBuiltin="1"');
        if (cs.iLevel !== undefined) attrs.push(`iLevel="${cs.iLevel}"`);
        if (cs.hidden) attrs.push('hidden="1"');
        csParts.push(`<cellStyle ${attrs.join(" ")}/>`);
      }
      csParts.push("</cellStyles>");
      p.push(csParts.join(""));
    } else {
      p.push(
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
      );
    }

    // dxfs
    if (this.dxfs.length > 0) {
      p.push(`<dxfs count="${this.dxfs.length}">`);
      for (const dxf of this.dxfs) {
        const dParts: string[] = [];
        if (dxf.font) dParts.push(`<font>${this.fontXmlStr(dxf.font)}</font>`);
        if (dxf.fill) {
          const bgColor = dxf.fill.color ? `<bgColor rgb="FF${dxf.fill.color}"/>` : "";
          const patAttrs = attrs({ patternType: dxf.fill.patternType ?? "solid" });
          dParts.push(`<fill><patternFill${patAttrs}>${bgColor}</patternFill></fill>`);
        }
        if (dxf.numFmt) dParts.push(`<numFmt formatCode="${escapeXml(dxf.numFmt)}"/>`);
        if (dxf.border) dParts.push(`<border>${this.borderXmlStr(dxf.border)}</border>`);
        if (dParts.length > 0) {
          p.push(`<dxf>${dParts.join("")}</dxf>`);
        } else {
          p.push("<dxf/>");
        }
      }
      p.push("</dxfs>");
    } else {
      p.push('<dxfs count="0"/>');
    }
    // tableStyles (CT_TableStyles)
    if (this.tableStyles && this.tableStyles.length > 0) {
      const tsParts: string[] = [
        `<tableStyles count="${this.tableStyles.length}" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16">`,
      ];
      for (const ts of this.tableStyles) {
        const tsAttrs: string[] = [`name="${escapeXml(ts.name)}"`];
        if (ts.pivot) tsAttrs.push('pivot="1"');
        if (ts.elements && ts.elements.length > 0) {
          tsParts.push(`<tableStyle ${tsAttrs.join(" ")}>`);
          for (const el of ts.elements) {
            const elAttrs: string[] = [`type="${el.type}"`];
            if (el.dxfId !== undefined) elAttrs.push(`dxfId="${el.dxfId}"`);
            if (el.button) elAttrs.push('button="1"');
            tsParts.push(`<tableStyleElement ${elAttrs.join(" ")}/>`);
          }
          tsParts.push("</tableStyle>");
        } else {
          tsParts.push(`<tableStyle ${tsAttrs.join(" ")}/>`);
        }
      }
      tsParts.push("</tableStyles>");
      p.push(tsParts.join(""));
    } else {
      p.push(
        '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>',
      );
    }

    // colors (optional color palette)
    if (this.colors) {
      const c = this.colors;
      const colorParts: string[] = ["<colors>"];
      if (c.indexedColors && c.indexedColors.length > 0) {
        colorParts.push("<indexedColors>");
        for (const ic of c.indexedColors) {
          colorParts.push(`<rgbColor rgb="${ic.rgb}"/>`);
        }
        colorParts.push("</indexedColors>");
      }
      if (c.mruColors && c.mruColors.length > 0) {
        colorParts.push("<mruColors>");
        for (const mc of c.mruColors) {
          colorParts.push(`<color rgb="FF${mc}"/>`);
        }
        colorParts.push("</mruColors>");
      }
      colorParts.push("</colors>");
      p.push(colorParts.join(""));
    }

    // extLst — style sheet extensions
    if (this.styleExtensions && this.styleExtensions.length > 0) {
      const extParts: string[] = ["<extLst>"];
      for (const ext of this.styleExtensions) {
        if (ext.content) {
          extParts.push(`<ext uri="${ext.uri}">${ext.content}</ext>`);
        } else {
          extParts.push(`<ext uri="${ext.uri}"/>`);
        }
      }
      extParts.push("</extLst>");
      p.push(extParts.join(""));
    } else {
      p.push("<extLst/>");
    }

    p.push("</styleSheet>");
    return p.join("");
  }

  private fontXmlStr(f: FontOptions): string {
    const parts: string[] = [];
    if (f.bold) parts.push("<b/>");
    if (f.italic) parts.push("<i/>");
    if (f.underline) parts.push("<u/>");
    if (f.strike) parts.push("<strike/>");
    if (f.outline) parts.push("<outline/>");
    if (f.shadow) parts.push("<shadow/>");
    if (f.condense) parts.push("<condense/>");
    if (f.extend) parts.push("<extend/>");
    if (f.size) parts.push(`<sz val="${f.size}"/>`);
    if (f.color) parts.push(`<color rgb="FF${f.color}"/>`);
    if (f.font) parts.push(`<name val="${escapeXml(f.font)}"/>`);
    if (f.charset !== undefined) parts.push(`<charset val="${f.charset}"/>`);
    if (f.family !== undefined) parts.push(`<family val="${f.family}"/>`);
    if (f.vertAlign) parts.push(`<vertAlign val="${f.vertAlign}"/>`);
    if (f.scheme) parts.push(`<scheme val="${f.scheme}"/>`);
    return parts.join("");
  }

  private borderXmlStr(b: BorderSideOptions): string {
    const parts: string[] = [];
    const renderSide = (name: string, opts: BorderOptions | undefined, required = true) => {
      if (opts && opts.style && opts.style !== "none") {
        const colorStr = opts.color ? `<color rgb="FF${opts.color}"/>` : "";
        parts.push(`<${name} style="${opts.style}">${colorStr}</${name}>`);
      } else if (required) {
        parts.push(`<${name}/>`);
      }
    };
    for (const side of [
      "left",
      "right",
      "top",
      "bottom",
      "diagonal",
      "vertical",
      "horizontal",
    ] as const) {
      renderSide(side, b[side] as BorderOptions | undefined);
    }
    // start/end not in transitional XSD — only emit when styled
    renderSide("start", b.start, false);
    renderSide("end", b.end, false);
    return parts.join("");
  }

  private alignmentXmlStr(a: AlignmentOptions): string {
    const aAttrs: Record<string, string | number | boolean | undefined> = {};
    if (a.horizontal) aAttrs.horizontal = a.horizontal;
    if (a.vertical) aAttrs.vertical = a.vertical;
    if (a.wrapText) aAttrs.wrapText = 1;
    if (a.textRotation !== undefined) aAttrs.textRotation = a.textRotation;
    if (a.indent !== undefined) aAttrs.indent = a.indent;
    if (a.relativeIndent !== undefined) aAttrs.relativeIndent = a.relativeIndent;
    if (a.justifyLastLine) aAttrs.justifyLastLine = 1;
    if (a.shrinkToFit) aAttrs.shrinkToFit = 1;
    if (a.readingOrder !== undefined) aAttrs.readingOrder = a.readingOrder;
    return `<alignment${attrs(aAttrs)}/>`;
  }

  private protectionXmlStr(pr: CellProtectionOptions): string {
    const prAttrs: Record<string, string | number | boolean | undefined> = {};
    if (pr.locked !== undefined) prAttrs.locked = pr.locked ? 1 : 0;
    if (pr.hidden !== undefined) prAttrs.hidden = pr.hidden ? 1 : 0;
    return `<protection${attrs(prAttrs)}/>`;
  }
}

// ── Descriptor Types ──

export interface StylesDocOptions {
  /** The Styles accumulator instance (for stringify). */
  styles: Styles;
}

// ── Descriptor ──

export const stylesDesc: CustomDescriptor<StylesDocOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return opts.styles.serialize();
  },

  parse(el, _ctx) {
    const result: Record<string, unknown> = {};

    // numFmts
    const numFmtsEl = findChild(el, "numFmts");
    if (numFmtsEl) {
      const numFmts: Record<string, number> = {};
      for (const nf of numFmtsEl.elements ?? []) {
        if (nf.name !== "numFmt") continue;
        const id = attrNum(nf, "numFmtId");
        const code = attr(nf, "formatCode");
        if (id !== undefined && code) numFmts[code] = id;
      }
      result.customNumFmts = numFmts;
    }

    // fonts
    const fontsEl = findChild(el, "fonts");
    if (fontsEl) {
      const fonts: FontOptions[] = [];
      for (const f of fontsEl.elements ?? []) {
        if (f.name !== "font") continue;
        fonts.push(parseFont(f));
      }
      result.fonts = fonts;
    }

    // fills
    const fillsEl = findChild(el, "fills");
    if (fillsEl) {
      const fills: FillOptions[] = [];
      for (const f of fillsEl.elements ?? []) {
        if (f.name !== "fill") continue;
        fills.push(parseFill(f));
      }
      result.fills = fills;
    }

    // borders
    const bordersEl = findChild(el, "borders");
    if (bordersEl) {
      const borders: BorderSideOptions[] = [];
      for (const b of bordersEl.elements ?? []) {
        if (b.name !== "border") continue;
        borders.push(parseBorder(b));
      }
      result.borders = borders;
    }

    // cellStyleXfs
    const cellStyleXfsEl = findChild(el, "cellStyleXfs");
    if (cellStyleXfsEl) {
      const xfs: Record<string, unknown>[] = [];
      for (const xf of cellStyleXfsEl.elements ?? []) {
        if (xf.name !== "xf") continue;
        const style: Record<string, unknown> = {};
        const fontId = attrNum(xf, "fontId");
        const fillId = attrNum(xf, "fillId");
        const borderId = attrNum(xf, "borderId");
        const numFmtId = attrNum(xf, "numFmtId");
        if (fontId !== undefined) style.fontIdx = fontId;
        if (fillId !== undefined) style.fillIdx = fillId;
        if (borderId !== undefined) style.borderIdx = borderId;
        if (numFmtId !== undefined) style.numFmtIdx = numFmtId;
        xfs.push(style);
      }
      result.cellStyleXfs = xfs;
    }

    // cellXfs
    const cellXfsEl = findChild(el, "cellXfs");
    if (cellXfsEl) {
      const xfs: StyleOptions[] = [];
      for (const xf of cellXfsEl.elements ?? []) {
        if (xf.name !== "xf") continue;
        const fontId = attrNum(xf, "fontId") ?? 0;
        const fillId = attrNum(xf, "fillId") ?? 0;
        const borderId = attrNum(xf, "borderId") ?? 0;
        const numFmtId = attrNum(xf, "numFmtId") ?? 0;

        const alignmentEl = findChild(xf, "alignment");
        const alignment = alignmentEl ? parseAlignment(alignmentEl) : undefined;

        const protectionEl = findChild(xf, "protection");
        const protection = protectionEl ? parseProtection(protectionEl) : undefined;

        const style: Record<string, unknown> = {};
        if (fontId > 0) style.fontIdx = fontId;
        if (fillId > 0) style.fillIdx = fillId;
        if (borderId > 0) style.borderIdx = borderId;
        if (numFmtId > 0) style.numFmtIdx = numFmtId;
        if (alignment) style.alignment = alignment;
        if (protection) style.protection = protection;
        if (attr(xf, "quotePrefix") === "1") style.quotePrefix = true;
        if (attr(xf, "pivotButton") === "1") style.pivotButton = true;

        xfs.push(style as StyleOptions);
      }
      result.cellXfs = xfs;
    }

    // cellStyles
    const cellStylesEl = findChild(el, "cellStyles");
    if (cellStylesEl) {
      const styles: CustomCellStyleOptions[] = [];
      for (const cs of cellStylesEl.elements ?? []) {
        if (cs.name !== "cellStyle") continue;
        const style: Record<string, unknown> = {};
        if (attr(cs, "name")) style.name = attr(cs, "name");
        const xfId = attrNum(cs, "xfId");
        if (xfId !== undefined) style.xfId = xfId;
        const builtinId = attrNum(cs, "builtinId");
        if (builtinId !== undefined) style.builtinId = builtinId;
        if (attr(cs, "customBuiltin") === "1") style.customBuiltin = true;
        if (attr(cs, "hidden") === "1") style.hidden = true;
        const iLevel = attrNum(cs, "iLevel");
        if (iLevel !== undefined) style.iLevel = iLevel;
        styles.push(style as unknown as CustomCellStyleOptions);
      }
      result.customCellStyles = styles;
    }

    // dxfs
    const dxfsEl = findChild(el, "dxfs");
    if (dxfsEl) {
      const dxfs: DxfOptions[] = [];
      for (const dxf of dxfsEl.elements ?? []) {
        if (dxf.name !== "dxf") continue;
        const d: Record<string, unknown> = {};
        const fontEl = findChild(dxf, "font");
        if (fontEl) d.font = parseFont(fontEl);
        const fillEl = findChild(dxf, "fill");
        if (fillEl) d.fill = parseFill(fillEl);
        const borderEl = findChild(dxf, "border");
        if (borderEl) d.border = parseBorder(borderEl);
        const numFmtEl = findChild(dxf, "numFmt");
        if (numFmtEl && attr(numFmtEl, "formatCode")) d.numFmt = attr(numFmtEl, "formatCode");
        dxfs.push(d as DxfOptions);
      }
      result.dxfs = dxfs;
    }

    // tableStyles
    const tableStylesEl = findChild(el, "tableStyles");
    if (tableStylesEl?.attributes) {
      const ts: Record<string, unknown> = {};
      if (attr(tableStylesEl, "count") !== undefined)
        ts.count = attrNum(tableStylesEl, "count") ?? 0;
      if (attr(tableStylesEl, "defaultTableStyle"))
        ts.defaultTableStyle = attr(tableStylesEl, "defaultTableStyle");
      if (attr(tableStylesEl, "defaultPivotStyle"))
        ts.defaultPivotStyle = attr(tableStylesEl, "defaultPivotStyle");
      const customStyles: CustomTableStyleOptions[] = [];
      for (const tse of tableStylesEl.elements ?? []) {
        if (tse.name !== "tableStyle") continue;
        const style: Record<string, unknown> = {};
        if (attr(tse, "name")) style.name = attr(tse, "name");
        if (attr(tse, "pivot") === "1") style.pivot = true;
        const elements: TableStyleElementOptions[] = [];
        for (const tsee of tse.elements ?? []) {
          if (tsee.name !== "tableStyleElement") continue;
          const elOpts: Record<string, unknown> = {};
          if (attr(tsee, "type")) elOpts.type = attr(tsee, "type") as TableStyleElementType;
          const dxfId = attrNum(tsee, "dxfId");
          if (dxfId !== undefined) elOpts.dxfId = dxfId;
          if (attr(tsee, "button") === "1") elOpts.button = true;
          elements.push(elOpts as unknown as TableStyleElementOptions);
        }
        if (elements.length > 0) style.elements = elements;
        customStyles.push(style as unknown as CustomTableStyleOptions);
      }
      if (customStyles.length > 0) ts.tableStyles = customStyles;
      result.tableStylesInfo = ts;
    }

    // colors
    const colorsEl = findChild(el, "colors");
    if (colorsEl) {
      const colors: Record<string, unknown> = {};
      const icEl = findChild(colorsEl, "indexedColors");
      if (icEl) {
        const indexed: IndexedColorOptions[] = [];
        for (const rgb of icEl.elements ?? []) {
          if (rgb.name === "rgbColor" && attr(rgb, "rgb")) {
            indexed.push({ rgb: attr(rgb, "rgb")! });
          }
        }
        colors.indexedColors = indexed;
      }
      const mruEl = findChild(colorsEl, "mruColors");
      if (mruEl) {
        const mru: string[] = [];
        for (const c of mruEl.elements ?? []) {
          if (c.name === "color") {
            const rgb = attr(c, "rgb");
            if (rgb) mru.push(rgb.length === 8 ? rgb.slice(2) : rgb);
          }
        }
        colors.mruColors = mru;
      }
      result.colors = colors;
    }

    // styleExtensions (extLst)
    const extLstEl = findChild(el, "extLst");
    if (extLstEl) {
      const exts: StyleExtensionOptions[] = [];
      for (const ext of extLstEl.elements ?? []) {
        if (ext.name !== "ext") continue;
        const uri = attr(ext, "uri");
        if (uri) {
          // Serialize child elements back as raw XML content
          const content = (ext.elements ?? [])
            .map((e) => {
              // Simple reconstruction for extension content
              return JSON.stringify(e);
            })
            .join("");
          exts.push({ uri, content: content || undefined });
        }
      }
      result.styleExtensions = exts;
    }

    return result as unknown as StylesDocOptions;
  },
};

// ── Parse helpers ──

function parseFont(el: XmlElement): FontOptions {
  const result: Record<string, unknown> = {};
  for (const child of el.elements ?? []) {
    switch (child.name) {
      case "b":
        result.bold = true;
        break;
      case "i":
        result.italic = true;
        break;
      case "u":
        result.underline = true;
        break;
      case "strike":
        result.strike = true;
        break;
      case "outline":
        result.outline = true;
        break;
      case "shadow":
        result.shadow = true;
        break;
      case "condense":
        result.condense = true;
        break;
      case "extend":
        result.extend = true;
        break;
      case "sz":
        result.size = attrNum(child, "val");
        break;
      case "color":
        result.color = parseColorHex(child);
        break;
      case "name":
        result.font = attr(child, "val") ?? undefined;
        break;
      case "charset":
        result.charset = attrNum(child, "val");
        break;
      case "family":
        result.family = attrNum(child, "val");
        break;
      case "vertAlign":
        result.vertAlign = (attr(child, "val") as FontOptions["vertAlign"]) ?? undefined;
        break;
      case "scheme":
        result.scheme = (attr(child, "val") as FontOptions["scheme"]) ?? undefined;
        break;
    }
  }
  return result as unknown as FontOptions;
}

function parseFill(el: XmlElement): FillOptions {
  const patternFill = findChild(el, "patternFill");
  if (patternFill) {
    const result: Record<string, unknown> = {};
    result.patternType = attr(patternFill, "patternType") ?? undefined;
    const fg = findChild(patternFill, "fgColor");
    if (fg) result.color = parseColorHex(fg);
    const bg = findChild(patternFill, "bgColor");
    if (bg) result.bgColor = parseColorHex(bg);
    const indexed = fg ? attrNum(fg, "indexed") : undefined;
    if (indexed !== undefined) result.colorIndexed = indexed;
    return result as unknown as FillOptions;
  }

  const gradientFill = findChild(el, "gradientFill");
  if (gradientFill) {
    const result: Record<string, unknown> = { type: "gradient" };
    const gType = attr(gradientFill, "type");
    if (gType) result.gradientType = gType as FillOptions["gradientType"];
    const degree = attrNum(gradientFill, "degree");
    if (degree !== undefined) result.gradientDegree = degree;
    const stops: { position: number; color: string }[] = [];
    for (const s of gradientFill.elements ?? []) {
      if (s.name !== "stop") continue;
      const pos = attrNum(s, "position");
      const color = findChild(s, "color");
      if (pos !== undefined && color) {
        stops.push({ position: pos, color: parseColorHex(color) ?? "" });
      }
    }
    if (stops.length > 0) result.stops = stops;
    return result as unknown as FillOptions;
  }

  return {} as unknown as FillOptions;
}

function parseBorder(el: XmlElement): BorderSideOptions {
  const result: Record<string, unknown> = {};
  if (attr(el, "diagonalUp") === "1") result.diagonalUp = true;
  if (attr(el, "diagonalDown") === "1") result.diagonalDown = true;

  for (const side of [
    "left",
    "right",
    "top",
    "bottom",
    "diagonal",
    "start",
    "end",
    "vertical",
    "horizontal",
  ] as const) {
    const sideEl = findChild(el, side);
    if (sideEl) {
      const opts: Record<string, unknown> = {};
      const style = attr(sideEl, "style");
      if (style) opts.style = style as BorderOptions["style"];
      const color = findChild(sideEl, "color");
      if (color) opts.color = parseColorHex(color);
      if (Object.keys(opts).length > 0) result[side] = opts;
    }
  }

  return result as unknown as BorderSideOptions;
}

function parseAlignment(el: XmlElement): AlignmentOptions {
  const result: Record<string, unknown> = {};
  const h = attr(el, "horizontal");
  if (h) result.horizontal = h as AlignmentOptions["horizontal"];
  const v = attr(el, "vertical");
  if (v) result.vertical = v as AlignmentOptions["vertical"];
  if (attr(el, "wrapText") === "1") result.wrapText = true;
  const rotation = attrNum(el, "textRotation");
  if (rotation !== undefined) result.textRotation = rotation;
  const indent = attrNum(el, "indent");
  if (indent !== undefined) result.indent = indent;
  return result as unknown as AlignmentOptions;
}

function parseProtection(el: XmlElement): CellProtectionOptions {
  const result: Record<string, unknown> = {};
  const locked = attr(el, "locked");
  if (locked !== undefined) result.locked = locked !== "0";
  const hidden = attr(el, "hidden");
  if (hidden !== undefined) result.hidden = hidden !== "0";
  return result as unknown as CellProtectionOptions;
}

function parseColorHex(el: XmlElement): string | undefined {
  const rgb = attr(el, "rgb");
  if (rgb) {
    // Strip alpha prefix if present (FF000000 → 000000)
    return rgb.length === 8 ? rgb.slice(2) : rgb;
  }
  return undefined;
}
