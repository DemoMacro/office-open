/**
 * Styles component — generates xl/styles.xml.
 *
 * XLSX uses an index-based style system: cells reference style entries
 * via the `s` attribute, which is an index into `cellXfs`.
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { attrs, escapeXml } from "@office-open/xml";

// ── Sub-style option interfaces ──

export interface FontOptions {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly strike?: boolean;
  readonly size?: number;
  readonly color?: string;
  readonly fontName?: string;
  /** Character set (CT_Font/charset @val) */
  readonly charset?: number;
  /** Font family (CT_Font/family @val) */
  readonly family?: number;
  /** Condense (macOS, CT_Font/condense) */
  readonly condense?: boolean;
  /** Extend (macOS, CT_Font/extend) */
  readonly extend?: boolean;
  /** Vertical alignment: superscript/subscript (CT_Font/vertAlign @val) */
  readonly vertAlign?: "superscript" | "subscript" | "baseline";
  /** Font scheme (CT_Font/scheme @val) */
  readonly scheme?: "major" | "minor" | "none";
  /** Font shadow (CT_Font/shadow) */
  readonly shadow?: boolean;
  /** Font outline (CT_Font/outline) */
  readonly outline?: boolean;
}

/** Gradient stop (CT_GradientStop) */
export interface GradientStopOptions {
  /** Position (0.0–1.0) */
  readonly position: number;
  /** RGB color hex without alpha, e.g. "FF0000" */
  readonly color: string;
}

export interface FillOptions {
  readonly type?: "solid" | "pattern" | "gradient";
  readonly color?: string;
  readonly patternType?: string;
  /** Background color for pattern fill (CT_PatternFill/bgColor) */
  readonly bgColor?: string;
  /** Gradient stops (CT_GradientFill/stop) */
  readonly stops?: readonly GradientStopOptions[];
  /** Gradient type (CT_GradientFill @type) */
  readonly gradientType?: "linear" | "path";
  /** Gradient degree for linear (CT_GradientFill @degree) */
  readonly gradientDegree?: number;
  /** Gradient left position for path (CT_GradientFill @left) */
  readonly gradientLeft?: number;
  /** Gradient right position for path (CT_GradientFill @right) */
  readonly gradientRight?: number;
  /** Gradient top position for path (CT_GradientFill @top) */
  readonly gradientTop?: number;
  /** Gradient bottom position for path (CT_GradientFill @bottom) */
  readonly gradientBottom?: number;
}

export interface BorderOptions {
  readonly style?: "thin" | "medium" | "thick" | "dotted" | "dashed" | "hair" | "none";
  readonly color?: string;
  /** Diagonal up (CT_Border @diagonalUp) — on the parent border element */
  readonly diagonalUp?: boolean;
  /** Diagonal down (CT_Border @diagonalDown) — on the parent border element */
  readonly diagonalDown?: boolean;
}

export interface BorderSideOptions {
  readonly top?: BorderOptions;
  readonly bottom?: BorderOptions;
  readonly left?: BorderOptions;
  readonly right?: BorderOptions;
  readonly diagonal?: BorderOptions;
  /** Leading edge border (CT_Border/start, for RTL support) */
  readonly start?: BorderOptions;
  /** Trailing edge border (CT_Border/end, for RTL support) */
  readonly end?: BorderOptions;
  /** Vertical inner border (CT_Border/vertical, for cell range borders) */
  readonly vertical?: BorderOptions;
  /** Horizontal inner border (CT_Border/horizontal, for cell range borders) */
  readonly horizontal?: BorderOptions;
}

export interface AlignmentOptions {
  readonly horizontal?: "left" | "center" | "right" | "fill" | "justify";
  readonly vertical?: "top" | "center" | "bottom";
  readonly wrapText?: boolean;
  readonly textRotation?: number;
  readonly indent?: number;
  /** Relative indent (CT_CellAlignment @relativeIndent) */
  readonly relativeIndent?: number;
  /** Justify last line (CT_CellAlignment @justifyLastLine) */
  readonly justifyLastLine?: boolean;
  /** Shrink to fit (CT_CellAlignment @shrinkToFit) */
  readonly shrinkToFit?: boolean;
  /** Reading order (CT_CellAlignment @readingOrder) */
  readonly readingOrder?: number;
}

export interface StyleOptions {
  readonly font?: FontOptions;
  readonly fill?: FillOptions;
  readonly border?: BorderSideOptions;
  readonly numFmt?: string;
  readonly alignment?: AlignmentOptions;
  /** Quote prefix (CT_Xf @quotePrefix) */
  readonly quotePrefix?: boolean;
  /** Pivot button (CT_Xf @pivotButton) */
  readonly pivotButton?: boolean;
  /** Cell protection (CT_CellProtection) */
  readonly protection?: CellProtectionOptions;
}

/** Cell-level protection settings (CT_CellProtection) */
export interface CellProtectionOptions {
  /** Cell is locked (CT_CellProtection @locked) */
  readonly locked?: boolean;
  /** Cell formula is hidden (CT_CellProtection @hidden) */
  readonly hidden?: boolean;
}

/** Indexed color entry (CT_RgbColor) */
export interface IndexedColorOptions {
  /** RGB hex value, e.g. "FF000000" */
  readonly rgb: string;
}

/** Colors palette (CT_Colors) */
export interface ColorsOptions {
  /** Indexed color palette (CT_IndexedColors) */
  readonly indexedColors?: readonly IndexedColorOptions[];
  /** Most recently used colors (CT_MRUColors) */
  readonly mruColors?: readonly string[];
}

/** Differential format — used by conditional formatting to specify what changes. */
export interface DxfOptions {
  readonly font?: FontOptions;
  readonly fill?: FillOptions;
  readonly border?: BorderSideOptions;
  readonly numFmt?: string;
}

// ── Style key helpers for deduplication ──

function fontKey(f: FontOptions): string {
  return `b${f.bold ? 1 : 0}i${f.italic ? 1 : 0}u${f.underline ? 1 : 0}s${f.strike ? 1 : 0}z${f.size ?? 0}c${f.color ?? ""}n${f.fontName ?? ""}cs${f.charset ?? ""}fm${f.family ?? ""}co${f.condense ? 1 : 0}ex${f.extend ? 1 : 0}va${f.vertAlign ?? ""}sc${f.scheme ?? ""}sh${f.shadow ? 1 : 0}ol${f.outline ? 1 : 0}`;
}

function fillKey(f: FillOptions): string {
  return `t${f.type ?? ""}c${f.color ?? ""}p${f.patternType ?? ""}bg${f.bgColor ?? ""}g${f.stops?.map((s) => `${s.position}_${s.color}`).join("|") ?? ""}`;
}

function borderKey(b: BorderSideOptions): string {
  const sk = (o?: BorderOptions) => `${o?.style ?? ""}_${o?.color ?? ""}`;
  return `t${sk(b.top)}b${sk(b.bottom)}l${sk(b.left)}r${sk(b.right)}d${sk(b.diagonal)}st${sk(b.start)}en${sk(b.end)}v${sk(b.vertical)}h${sk(b.horizontal)}`;
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
  readonly type: TableStyleElementType;
  /** Differential format index (dxf) */
  readonly dxfId?: number;
  /** Button style (for pivot tables) */
  readonly button?: boolean;
}

/** Custom table/pivot table style (CT_TableStyle). */
export interface CustomTableStyleOptions {
  /** Style name (must be unique) */
  readonly name: string;
  /** Pivot style (vs table style) */
  readonly pivot?: boolean;
  /** Table style elements */
  readonly elements?: readonly TableStyleElementOptions[];
}

export class Styles extends BaseXmlComponent {
  private readonly fonts: FontOptions[] = [
    { size: 11, fontName: "Calibri" }, // default font (index 0)
  ];
  private readonly fontKeys = new Map<string, number>();

  private readonly fills: FillOptions[] = [
    { patternType: "none" }, // default fill (index 0)
    { patternType: "gray125" }, // required fill (index 1)
  ];
  private readonly fillKeys = new Map<string, number>();

  private readonly borders: BorderSideOptions[] = [
    {}, // default empty border (index 0)
  ];
  private readonly borderKeys = new Map<string, number>();

  private readonly customNumFmts = new Map<string, number>();
  private nextCustomNumFmtId = 164; // custom numFmts start at 164

  private readonly cellXfs: Array<{
    readonly fontId: number;
    readonly fillId: number;
    readonly borderId: number;
    readonly numFmtId: number;
    readonly alignment?: AlignmentOptions;
    readonly quotePrefix?: boolean;
    readonly pivotButton?: boolean;
    readonly protection?: CellProtectionOptions;
  }> = [
    { fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 }, // default xf (index 0)
  ];
  private readonly cellXfKeys = new Map<string, number>();

  private readonly dxfs: DxfOptions[] = [];

  private colors?: ColorsOptions;
  private tableStyles?: readonly CustomTableStyleOptions[];

  public constructor() {
    super("styleSheet");

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

  public setTableStyles(styles: readonly CustomTableStyleOptions[]): void {
    this.tableStyles = styles;
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
    readonly fontId: number;
    readonly fillId: number;
    readonly borderId: number;
    readonly numFmtId: number;
    readonly alignment?: AlignmentOptions;
    readonly quotePrefix?: boolean;
    readonly pivotButton?: boolean;
    readonly protection?: CellProtectionOptions;
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
  public override toXml(_context: Context): string {
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
        const fgColor = f.color ? `<fgColor rgb="FF${f.color}"/>` : "";
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
      p.push(`<border>${this.borderXmlStr(b)}</border>`);
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

      const alignStr = xf.alignment ? this.alignmentXmlStr(xf.alignment) : "";
      const protStr = xf.protection ? this.protectionXmlStr(xf.protection) : "";
      const inner = alignStr + protStr;
      p.push(inner ? `<xf${attrs(xAttrs)}>${inner}</xf>` : `<xf${attrs(xAttrs)}/>`);
    }
    p.push("</cellXfs>");

    // cellStyles
    p.push('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>');

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

    p.push("<extLst/>");

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
    if (f.fontName) parts.push(`<name val="${escapeXml(f.fontName)}"/>`);
    if (f.charset !== undefined) parts.push(`<charset val="${f.charset}"/>`);
    if (f.family !== undefined) parts.push(`<family val="${f.family}"/>`);
    if (f.vertAlign) parts.push(`<vertAlign val="${f.vertAlign}"/>`);
    if (f.scheme) parts.push(`<scheme val="${f.scheme}"/>`);
    return parts.join("");
  }

  private borderXmlStr(b: BorderSideOptions): string {
    const parts: string[] = [];
    const renderSide = (name: string, opts: BorderOptions | undefined) => {
      if (opts && opts.style && opts.style !== "none") {
        const colorStr = opts.color ? `<color rgb="FF${opts.color}"/>` : "";
        parts.push(`<${name} style="${opts.style}">${colorStr}</${name}>`);
      } else {
        parts.push(`<${name}/>`);
      }
    };
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
      renderSide(side, b[side] as BorderOptions | undefined);
    }
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
