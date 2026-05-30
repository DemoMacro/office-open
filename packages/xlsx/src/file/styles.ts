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
}

export interface FillOptions {
  readonly type?: "solid" | "pattern";
  readonly color?: string;
  readonly patternType?: string;
}

export interface BorderOptions {
  readonly style?: "thin" | "medium" | "thick" | "dotted" | "dashed" | "hair" | "none";
  readonly color?: string;
}

export interface BorderSideOptions {
  readonly top?: BorderOptions;
  readonly bottom?: BorderOptions;
  readonly left?: BorderOptions;
  readonly right?: BorderOptions;
  readonly diagonal?: BorderOptions;
}

export interface AlignmentOptions {
  readonly horizontal?: "left" | "center" | "right" | "fill" | "justify";
  readonly vertical?: "top" | "center" | "bottom";
  readonly wrapText?: boolean;
  readonly textRotation?: number;
  readonly indent?: number;
}

export interface StyleOptions {
  readonly font?: FontOptions;
  readonly fill?: FillOptions;
  readonly border?: BorderSideOptions;
  readonly numFmt?: string;
  readonly alignment?: AlignmentOptions;
}

// ── Style key helpers for deduplication ──

function fontKey(f: FontOptions): string {
  return `b${f.bold ? 1 : 0}i${f.italic ? 1 : 0}u${f.underline ? 1 : 0}s${f.strike ? 1 : 0}z${f.size ?? 0}c${f.color ?? ""}n${f.fontName ?? ""}`;
}

function fillKey(f: FillOptions): string {
  return `t${f.type ?? ""}c${f.color ?? ""}p${f.patternType ?? ""}`;
}

function borderKey(b: BorderSideOptions): string {
  const sk = (o?: BorderOptions) => `${o?.style ?? ""}_${o?.color ?? ""}`;
  return `t${sk(b.top)}b${sk(b.bottom)}l${sk(b.left)}r${sk(b.right)}d${sk(b.diagonal)}`;
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
  }> = [
    { fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 }, // default xf (index 0)
  ];
  private readonly cellXfKeys = new Map<string, number>();

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
    };

    const key = this.cellXfKey(xf);
    const existing = this.cellXfKeys.get(key);
    if (existing !== undefined) return existing;

    const idx = this.cellXfs.length;
    this.cellXfs.push(xf);
    this.cellXfKeys.set(key, idx);
    return idx;
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
  }): string {
    const a = xf.alignment;
    const ak = a
      ? `h${a.horizontal ?? ""}v${a.vertical ?? ""}w${a.wrapText ? 1 : 0}r${a.textRotation ?? ""}i${a.indent ?? ""}`
      : "";
    return `${xf.fontId}|${xf.fillId}|${xf.borderId}|${xf.numFmtId}|${ak}`;
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
      const patternAttrs = attrs({ patternType: f.patternType ?? "solid" });
      const fgColor = f.color ? `<fgColor rgb="FF${f.color}"/>` : "";
      p.push(
        fgColor
          ? `<fill><patternFill${patternAttrs}>${fgColor}</patternFill></fill>`
          : `<fill><patternFill${patternAttrs}/></fill>`,
      );
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

      const alignStr = xf.alignment ? this.alignmentXmlStr(xf.alignment) : "";
      p.push(alignStr ? `<xf${attrs(xAttrs)}>${alignStr}</xf>` : `<xf${attrs(xAttrs)}/>`);
    }
    p.push("</cellXfs>");

    // cellStyles
    p.push('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>');

    // dxfs and tableStyles (required by Excel)
    p.push('<dxfs count="0"/>');
    p.push(
      '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>',
    );
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
    if (f.size) parts.push(`<sz val="${f.size}"/>`);
    if (f.color) parts.push(`<color rgb="FF${f.color}"/>`);
    if (f.fontName) parts.push(`<name val="${escapeXml(f.fontName)}"/>`);
    return parts.join("");
  }

  private borderXmlStr(b: BorderSideOptions): string {
    const parts: string[] = [];
    for (const side of ["left", "right", "top", "bottom", "diagonal"] as const) {
      const opts = b[side] as BorderOptions | undefined;
      if (opts && opts.style && opts.style !== "none") {
        const colorStr = opts.color ? `<color rgb="FF${opts.color}"/>` : "";
        parts.push(`<${side} style="${opts.style}">${colorStr}</${side}>`);
      } else {
        parts.push(`<${side}/>`);
      }
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
    return `<alignment${attrs(aAttrs)}/>`;
  }
}
