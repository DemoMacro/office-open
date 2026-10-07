/**
 * Styles — the xl/styles.xml accumulator.
 *
 * XLSX uses an index-based style system: cells reference style entries
 * via the `s` attribute, which is an index into `cellXfs`. This class
 * registers and deduplicates fonts, fills, borders, numFmts, and cellXfs,
 * then serializes the lot via {@link Styles.serialize}.
 *
 * @module
 */
import { attrs, escapeXml } from "@office-open/xml";

import type {
  AlignmentOptions,
  BorderOptions,
  BorderSideOptions,
  CellProtectionOptions,
  CellStyleXfOptions,
  ColorsOptions,
  CustomCellStyleOptions,
  CellGradientStopOptions,
  DxfOptions,
  CellFillOptions,
  FontOptions,
  IndexedXfEntry,
  NumFmtEntry,
  TableStylesInfo,
  StyleExtensionOptions,
  StyleOptions,
  StylesState,
} from "./types";

// ── Style key helpers for deduplication ──

function fontKey(f: FontOptions): string {
  const underlineKey =
    f.underline === undefined
      ? ""
      : f.underline === false
        ? "0"
        : f.underline === true
          ? "1"
          : f.underline;
  return `b${f.bold ? 1 : 0}i${f.italic ? 1 : 0}u${underlineKey}s${f.strike ? 1 : 0}z${f.size ?? 0}c${f.color ?? ""}tc${f.themeColor ?? ""}ti${f.tint ?? ""}ix${f.colorIndexed ?? ""}a${f.autoColor ? 1 : 0}n${f.font ?? ""}cs${f.charset ?? ""}fm${f.family ?? ""}co${f.condense ? 1 : 0}ex${f.extend ? 1 : 0}va${f.vertAlign ?? ""}sc${f.scheme ?? ""}sh${f.shadow ? 1 : 0}ol${f.outline ? 1 : 0}`;
}

function fillKey(f: CellFillOptions): string {
  const stopKey = (s: CellGradientStopOptions) =>
    `${s.position}_${s.color ?? ""}_${s.themeColor ?? ""}_${s.tint ?? ""}`;
  return `t${f.type ?? ""}c${f.color ?? ""}tc${f.themeColor ?? ""}ti${f.tint ?? ""}ix${f.colorIndexed ?? ""}fa${f.fgAutoColor ? 1 : 0}p${f.patternType ?? ""}bg${f.bgColor ?? ""}bgtc${f.bgThemeColor ?? ""}bgti${f.bgTint ?? ""}bgix${f.bgColorIndexed ?? ""}bga${f.bgAutoColor ? 1 : 0}gt${f.gradientType ?? ""}gd${f.gradientDegree ?? ""}gl${f.gradientLeft ?? ""}gr${f.gradientRight ?? ""}gtp${f.gradientTop ?? ""}gb${f.gradientBottom ?? ""}s${f.stops?.map(stopKey).join("|") ?? ""}`;
}

function borderKey(b: BorderSideOptions): string {
  // Existence bit: an empty <vertical/> (side: {}) must not dedup against a
  // border without that side — adopted tables rebuild keys from raw entries.
  const sk = (o?: BorderOptions) =>
    `${o ? 1 : 0}_${o?.style ?? ""}_${o?.color ?? ""}_${o?.themeColor ?? ""}_${o?.tint ?? ""}_${o?.colorIndexed ?? ""}_${o?.autoColor ? 1 : 0}`;
  return `t${sk(b.top)}b${sk(b.bottom)}l${sk(b.left)}r${sk(b.right)}d${sk(b.diagonal)}du${b.diagonalUp ? 1 : 0}dd${b.diagonalDown ? 1 : 0}ol${b.outline === undefined ? "" : b.outline ? 1 : 0}st${sk(b.start)}en${sk(b.end)}v${sk(b.vertical)}h${sk(b.horizontal)}`;
}

function decimalAttr(value: number): string {
  const shortest = String(value);
  if (shortest.length < 17) return shortest;
  return value.toPrecision(17).replace(/0+$/, "").replace(/\.$/, "");
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

/** Reverse lookup: built-in number format id → format code. */
const BUILTIN_NUMFMT_BY_ID = new Map<number, string>(
  Object.entries(BUILTIN_NUMFMTS).map(([code, id]) => [id, code]),
);

/**
 * Format code for a built-in number format id (0-49/164+ reserved range),
 * or undefined for custom ids — those live in the parsed `<numFmts>` table.
 */
export function builtinNumFmtCode(id: number): string | undefined {
  return BUILTIN_NUMFMT_BY_ID.get(id);
}

/**
 * Internal cellStyleXfs storage — font/fill/border/numFmt indices after
 * re-registration against the accumulator, plus the verbatim CT_Xf flags.
 */
interface CellStyleXfEntry {
  fontId: number;
  fillId?: number;
  borderId?: number;
  numFmtId?: number;
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

interface FontsContainerState {
  knownFonts?: boolean;
}

export class Styles {
  private fonts: FontOptions[] = [
    { size: 11, font: "Calibri" }, // default font (index 0)
  ];
  private fontKeys = new Map<string, number>();
  private fontsContainer?: FontsContainerState;

  private fills: CellFillOptions[] = [
    { patternType: "none" }, // default fill (index 0)
    { patternType: "gray125" }, // required fill (index 1)
  ];
  private fillKeys = new Map<string, number>();

  private borders: BorderSideOptions[] = [
    {}, // default empty border (index 0)
  ];
  private borderKeys = new Map<string, number>();

  private customNumFmts: NumFmtEntry[] = [];
  private nextCustomNumFmtId = 164; // custom numFmts start at 164

  private cellXfs: Array<{
    fontId: number;
    fillId: number;
    borderId: number;
    numFmtId: number;
    fontIdDeclared?: boolean;
    fillIdDeclared?: boolean;
    borderIdDeclared?: boolean;
    numFmtIdDeclared?: boolean;
    /** cellStyleXfs reference from an adopted source xf; 0 = standalone. */
    xfId?: number;
    xfIdDeclared?: boolean;
    alignment?: AlignmentOptions;
    quotePrefix?: boolean;
    pivotButton?: boolean;
    applyProtection?: boolean;
    protection?: CellProtectionOptions;
    /** Explicit apply* flags from an adopted source xf; undefined = derive. */
    applyFont?: boolean;
    applyFill?: boolean;
    applyBorder?: boolean;
    applyNumberFormat?: boolean;
    applyAlignment?: boolean;
  }> = [
    { fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 }, // default xf (index 0)
  ];
  private cellXfKeys = new Map<string, number>();

  /** Named cell-style templates (cellStyleXfs); index 0 is the Normal style. */
  private cellStyleXfs: CellStyleXfEntry[] = [
    { fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 }, // Normal (index 0)
  ];

  private dxfs: DxfOptions[] = [];

  private colors?: ColorsOptions;
  private tableStyles?: TableStylesInfo;
  /** Custom cell styles (CT_CellStyles) */
  private customCellStyles?: CustomCellStyleOptions[];
  /** Style sheet extensions (CT_ExtensionList) */
  private styleExtensions?: StyleExtensionOptions[];
  /**
   * True once a parsed source table was adopted — switches optional
   * containers (dxfs, tableStyles) from the fresh-file convention (always
   * written, even empty) to emitting only what the source declared.
   */
  private roundTrip = false;
  /** True once setDxfs() ran — an explicit empty `<dxfs/>` still serializes. */
  private dxfsDeclared = false;
  /** True once an adopted table declared a (possibly empty) numFmts section. */
  private numFmtsDeclared = false;
  /** True once the source declared a (possibly empty) cellStyleXfs section. */
  private cellStyleXfsDeclared = false;

  public constructor() {
    // Pre-register default font/fill/border keys. These arrays are seeded
    // inline above, so index 0/1 always exist.
    this.fontKeys.set(fontKey(this.fonts[0]!), 0);
    this.fillKeys.set(fillKey(this.fills[0]!), 0);
    this.fillKeys.set(fillKey(this.fills[1]!), 1);
    this.borderKeys.set(borderKey(this.borders[0]!), 0);
    this.cellXfKeys.set(this.cellXfKey(this.cellXfs[0]!), 0);
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
   * Set the dxf list from parsed options. `[]` (the source declared an empty
   * `<dxfs/>` container) is distinct from never setting it (fresh document).
   */
  public setDxfs(list: DxfOptions[]): void {
    this.dxfs = [...list];
    this.dxfsDeclared = true;
  }

  /**
   * Set color palette (indexed colors and MRU colors).
   */
  public setColors(opts: ColorsOptions): void {
    this.colors = opts;
  }

  public setTableStyles(styles: TableStylesInfo): void {
    this.tableStyles = styles;
  }

  public setExtensions(extensions: StyleExtensionOptions[]): void {
    this.styleExtensions = extensions;
  }

  public setCustomCellStyles(styles: CustomCellStyleOptions[]): void {
    this.customCellStyles = styles;
  }

  /**
   * Set named cell-style templates (cellStyleXfs). Each entry's font/fill/
   * border/numFmt definitions are re-registered (deduplicated) so indices track
   * the rebuilt tables. Entries keep source order, keeping cellStyle.xfId
   * references stable (oldIndex === newIndex). Alignment/protection and the
   * applyXxx flags are preserved verbatim for named-style fidelity.
   */
  public setCellStyleXfs(entries: CellStyleXfOptions[]): void {
    this.cellStyleXfsDeclared = true;
    this.cellStyleXfs = entries.map((entry) => ({
      ...(entry.font !== undefined
        ? {
            fontId:
              this.adoptedDefinitionIndex(entry.font, this.fonts) ?? this.registerFont(entry.font),
          }
        : entry.fontIdRaw !== undefined
          ? { fontId: entry.fontIdRaw }
          : { fontId: 0 }),
      ...(entry.fillIdDeclared || entry.fill !== undefined
        ? {
            fillId:
              this.adoptedDefinitionIndex(entry.fill, this.fills) ?? this.registerFill(entry.fill),
          }
        : {}),
      ...(entry.borderIdDeclared || entry.border !== undefined
        ? {
            borderId:
              this.adoptedDefinitionIndex(entry.border, this.borders) ??
              this.registerBorder(entry.border),
          }
        : {}),
      ...(entry.numFmt !== undefined || entry.numFmtId !== undefined
        ? {
            numFmtId:
              entry.numFmt !== undefined ? this.registerNumFmt(entry.numFmt) : entry.numFmtId,
          }
        : {}),
      alignment: entry.alignment,
      protection: entry.protection,
      quotePrefix: entry.quotePrefix,
      pivotButton: entry.pivotButton,
      applyNumberFormat: entry.applyNumberFormat,
      applyFont: entry.applyFont,
      applyFill: entry.applyFill,
      applyBorder: entry.applyBorder,
      applyAlignment: entry.applyAlignment,
      applyProtection: entry.applyProtection,
    }));
  }

  private adoptedDefinitionIndex<T>(value: T | undefined, table: T[]): number | undefined {
    if (value === undefined) return undefined;
    const index = table.indexOf(value);
    return index >= 0 ? index : undefined;
  }

  /**
   * Expose internal state for descriptor-based XML generation.
   * The descriptor reads this snapshot to produce xl/styles.xml.
   */
  public toDescriptorOptions(): StylesState {
    return {
      customNumFmts: [...this.customNumFmts],
      fonts: [...this.fonts],
      fills: [...this.fills],
      borders: [...this.borders],
      cellXfs: [...this.cellXfs],
      dxfs: [...this.dxfs],
      colors: this.colors,
      tableStyles: this.tableStyles?.tableStyles,
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

  private registerFill(opts?: CellFillOptions): number {
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
    const declared = this.customNumFmts.find((entry) => entry.formatCode === fmt);
    if (declared) return declared.numFmtId;
    const builtin = BUILTIN_NUMFMTS[fmt];
    if (builtin !== undefined) return builtin;
    const id = this.nextCustomNumFmtId++;
    this.customNumFmts.push({ numFmtId: id, formatCode: fmt });
    return id;
  }

  private cellXfKey(xf: {
    fontId: number;
    fillId: number;
    borderId: number;
    numFmtId: number;
    xfId?: number;
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
    return `${xf.fontId}|${xf.fillId}|${xf.borderId}|${xf.numFmtId}|x${xf.xfId ?? ""}|${ak}|qp${xf.quotePrefix ? 1 : 0}|pb${xf.pivotButton ? 1 : 0}|${pk}`;
  }

  /**
   * Adopt a parsed source style table wholesale, replacing the fresh-file
   * defaults (the SDK's Stylesheet model: the table is a unit, cells carry
   * its indices). Fonts/fills/borders/cellXfs keep their source order, so raw
   * cell indices resolve exactly as they did in the source file; styles
   * registered afterwards dedup against the adopted entries and extend the
   * index space at the end.
   */
  public adopt(table: {
    fonts?: FontOptions[];
    fontsContainer?: FontsContainerState;
    fills?: CellFillOptions[];
    borders?: BorderSideOptions[];
    cellXfs?: IndexedXfEntry[];
    numFmts?: NumFmtEntry[];
  }): void {
    this.roundTrip = true;
    this.fonts = table.fonts ? [...table.fonts] : [this.fonts[0]!];
    this.fontKeys = new Map();
    this.fonts.forEach((f, i) => {
      const key = fontKey(f);
      if (!this.fontKeys.has(key)) this.fontKeys.set(key, i);
    });
    this.fontsContainer = table.fontsContainer;
    this.fills = table.fills
      ? [...table.fills]
      : [{ patternType: "none" }, { patternType: "gray125" }];
    this.fillKeys = new Map();
    this.fills.forEach((f, i) => {
      const key = fillKey(f);
      if (!this.fillKeys.has(key)) this.fillKeys.set(key, i);
    });
    this.borders = table.borders ? [...table.borders] : [{}];
    this.borderKeys = new Map();
    this.borders.forEach((b, i) => {
      const key = borderKey(b);
      if (!this.borderKeys.has(key)) this.borderKeys.set(key, i);
    });
    this.cellXfs = (table.cellXfs ?? [{ fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 }]).map(
      (e) => ({
        fontId: e.fontId ?? 0,
        fillId: e.fillId ?? 0,
        borderId: e.borderId ?? 0,
        numFmtId: e.numFmtId ?? 0,
        fontIdDeclared: e.fontIdDeclared,
        fillIdDeclared: e.fillIdDeclared,
        borderIdDeclared: e.borderIdDeclared,
        numFmtIdDeclared: e.numFmtIdDeclared,
        xfId: e.xfId,
        xfIdDeclared: e.xfIdDeclared,
        alignment: e.alignment,
        quotePrefix: e.quotePrefix,
        pivotButton: e.pivotButton,
        applyProtection: e.applyProtection,
        protection: e.protection,
        applyFont: e.applyFont,
        applyFill: e.applyFill,
        applyBorder: e.applyBorder,
        applyNumberFormat: e.applyNumberFormat,
        applyAlignment: e.applyAlignment,
      }),
    );
    this.cellXfKeys = new Map();
    this.cellXfs.forEach((xf, i) => {
      const key = this.cellXfKey(xf);
      if (!this.cellXfKeys.has(key)) this.cellXfKeys.set(key, i);
    });
    if (table.numFmts) {
      this.numFmtsDeclared = true;
      this.customNumFmts = table.numFmts.map((entry) => ({ ...entry }));
      for (const entry of this.customNumFmts) {
        if (entry.numFmtId >= this.nextCustomNumFmtId) {
          this.nextCustomNumFmtId = entry.numFmtId + 1;
        }
      }
    }
  }

  // ── XML generation ──

  /**
   * Zero-allocation fast path: directly concatenate XML string.
   * Bypasses the intermediate object tree entirely.
   */
  /** Serialize to xl/styles.xml content (without XML declaration). */
  public serialize(): string {
    const knownFonts = this.fontsContainer?.knownFonts;
    const p: string[] = [
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    ];

    // numFmts
    if (this.customNumFmts.length > 0) {
      p.push(`<numFmts count="${this.customNumFmts.length}">`);
      for (const entry of this.customNumFmts) {
        p.push(
          `<numFmt numFmtId="${entry.numFmtId}" formatCode="${escapeXml(entry.formatCode)}"/>`,
        );
      }
      p.push("</numFmts>");
    } else if (this.numFmtsDeclared) {
      p.push('<numFmts count="0"/>');
    }

    // Empty fonts adopt a required default; a source empty container is not
    // currently distinguishable in the compact Styles runtime.
    if (this.fonts.length > 0) {
      const fontAttrs =
        knownFonts === undefined
          ? ""
          : ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"' +
            ` x14ac:knownFonts="${knownFonts ? 1 : 0}"`;
      p.push(`<fonts count="${this.fonts.length}"${fontAttrs}>`);
      for (const f of this.fonts) {
        p.push(`<font>${this.fontXmlStr(f)}</font>`);
      }
      p.push("</fonts>");
    }

    // fills
    if (this.fills.length > 0) {
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
            .map((s) => {
              const channel =
                s.themeColor !== undefined
                  ? `theme="${s.themeColor}"`
                  : s.color
                    ? `rgb="FF${s.color}"`
                    : "";
              const tint =
                s.tintRaw !== undefined
                  ? ` tint="${s.tintRaw}"`
                  : s.tint !== undefined
                    ? ` tint="${decimalAttr(s.tint)}"`
                    : "";
              return `<stop position="${s.position}"><color ${channel}${tint}/></stop>`;
            })
            .join("");
          p.push(`<fill><gradientFill${attrs(gfAttrs)}>${stopParts}</gradientFill></fill>`);
        } else {
          const defaultPatternType =
            f.patternType ??
            (f.type === "solid" ? "solid" : f.patternTypeDeclared === false ? undefined : "solid");
          const patternAttrs =
            f.legacyPatternElement && f.legacyPatternType !== undefined
              ? ` patternType="${f.legacyPatternType}"`
              : attrs({ patternType: defaultPatternType });
          const fgChannel =
            f.fgLegacyColorType !== undefined
              ? `type="${f.fgLegacyColorType}" val="${f.fgLegacyColorValue ?? ""}"`
              : f.themeColor !== undefined
                ? `theme="${f.themeColor}"`
                : f.colorIndexed !== undefined
                  ? `indexed="${f.colorIndexed}"`
                  : f.color
                    ? `rgb="${f.fgColorRaw ?? `FF${f.color}`}"`
                    : f.fgAutoColor
                      ? 'auto="1"'
                      : "";
          const fgTint =
            f.tintRaw !== undefined
              ? ` tint="${f.tintRaw}"`
              : f.tint !== undefined
                ? ` tint="${decimalAttr(f.tint)}"`
                : "";
          const fgColor = fgChannel ? `<fgColor ${fgChannel}${fgTint}/>` : "";
          const bgChannel =
            f.bgLegacyColorType !== undefined
              ? `type="${f.bgLegacyColorType}" val="${f.bgLegacyColorValue ?? ""}"`
              : f.bgThemeColor !== undefined
                ? `theme="${f.bgThemeColor}"`
                : f.bgColorIndexed !== undefined
                  ? `indexed="${f.bgColorIndexed}"`
                  : f.bgColor
                    ? `rgb="${f.bgColorRaw ?? `FF${f.bgColor}`}"`
                    : f.bgAutoColor
                      ? 'auto="1"'
                      : "";
          const bgTint =
            f.bgTintRaw !== undefined
              ? ` tint="${f.bgTintRaw}"`
              : f.bgTint !== undefined
                ? ` tint="${decimalAttr(f.bgTint)}"`
                : "";
          const bgColor = bgChannel ? `<bgColor ${bgChannel}${bgTint}/>` : "";
          const colorContent = fgColor + bgColor;
          p.push(
            colorContent
              ? `<fill><${f.legacyPatternElement ? "pattern" : "patternFill"}${patternAttrs}>${colorContent}</${f.legacyPatternElement ? "pattern" : "patternFill"}></fill>`
              : `<fill><${f.legacyPatternElement ? "pattern" : "patternFill"}${patternAttrs}/></fill>`,
          );
        }
      }
      p.push("</fills>");
    }

    // borders
    if (this.borders.length > 0) {
      p.push(`<borders count="${this.borders.length}">`);
      for (const b of this.borders) {
        const bAttrs: string[] = [];
        if (b.diagonalUp !== undefined) bAttrs.push(`diagonalUp="${b.diagonalUp ? 1 : 0}"`);
        if (b.diagonalDown !== undefined) bAttrs.push(`diagonalDown="${b.diagonalDown ? 1 : 0}"`);
        if (b.outline !== undefined) bAttrs.push(`outline="${b.outline ? 1 : 0}"`);
        const bAttr = bAttrs.length ? ` ${bAttrs.join(" ")}` : "";
        const requiredSides = !this.roundTrip;
        p.push(`<border${bAttr}>${this.borderXmlStr(b, requiredSides)}</border>`);
      }
      p.push("</borders>");
    }

    // cellStyleXfs — named-style templates; applyXxx preserved verbatim (not derived)
    if (this.cellStyleXfs.length > 0 && (this.cellStyleXfsDeclared || !this.roundTrip)) {
      p.push(`<cellStyleXfs count="${this.cellStyleXfs.length}">`);
      for (const xf of this.cellStyleXfs) {
        const xAttrs: Record<string, string | number | boolean | undefined> = {
          ...(xf.numFmtId !== undefined ? { numFmtId: xf.numFmtId } : {}),
          fontId: xf.fontId,
          ...(xf.fillId !== undefined ? { fillId: xf.fillId } : {}),
          ...(xf.borderId !== undefined ? { borderId: xf.borderId } : {}),
        };
        if (xf.applyNumberFormat !== undefined)
          xAttrs.applyNumberFormat = xf.applyNumberFormat ? 1 : 0;
        if (xf.applyFont !== undefined) xAttrs.applyFont = xf.applyFont ? 1 : 0;
        if (xf.applyFill !== undefined) xAttrs.applyFill = xf.applyFill ? 1 : 0;
        if (xf.applyBorder !== undefined) xAttrs.applyBorder = xf.applyBorder ? 1 : 0;
        if (xf.applyAlignment !== undefined) xAttrs.applyAlignment = xf.applyAlignment ? 1 : 0;
        if (xf.applyProtection !== undefined) xAttrs.applyProtection = xf.applyProtection ? 1 : 0;
        if (xf.quotePrefix !== undefined) xAttrs.quotePrefix = xf.quotePrefix ? 1 : 0;
        if (xf.pivotButton !== undefined) xAttrs.pivotButton = xf.pivotButton ? 1 : 0;
        const alignStr = xf.alignment ? this.alignmentXmlStr(xf.alignment) : "";
        const protStr = xf.protection ? this.protectionXmlStr(xf.protection) : "";
        const inner = alignStr + protStr;
        p.push(inner ? `<xf${attrs(xAttrs)}>${inner}</xf>` : `<xf${attrs(xAttrs)}/>`);
      }
      p.push("</cellStyleXfs>");
    }

    // cellXfs
    if (this.cellXfs.length > 0) {
      p.push(`<cellXfs count="${this.cellXfs.length}">`);
      for (const xf of this.cellXfs) {
        const xAttrs: Record<string, string | number | boolean | undefined> = {
          numFmtId: xf.numFmtIdDeclared === false ? undefined : xf.numFmtId,
          fontId: xf.fontIdDeclared === false ? undefined : xf.fontId,
          fillId: xf.fillIdDeclared === false ? undefined : xf.fillId,
          borderId: xf.borderIdDeclared === false ? undefined : xf.borderId,
          xfId: xf.xfIdDeclared === false ? undefined : (xf.xfId ?? 0),
        };
        // Adopted xfs carry their source apply* flags verbatim; freshly
        // registered ones derive them from non-zero component ids.
        const derive = !this.roundTrip;
        const applyFont = xf.applyFont ?? (derive && xf.fontId > 0 ? true : undefined);
        const applyFill = xf.applyFill ?? (derive && xf.fillId > 0 ? true : undefined);
        const applyBorder = xf.applyBorder ?? (derive && xf.borderId > 0 ? true : undefined);
        const applyNumberFormat =
          xf.applyNumberFormat ?? (derive && xf.numFmtId > 0 ? true : undefined);
        const applyAlignment = xf.applyAlignment ?? (derive && xf.alignment ? true : undefined);
        const applyProtection = xf.applyProtection ?? (derive && xf.protection ? true : undefined);
        if (applyFont !== undefined) xAttrs.applyFont = applyFont ? 1 : 0;
        if (applyFill !== undefined) xAttrs.applyFill = applyFill ? 1 : 0;
        if (applyBorder !== undefined) xAttrs.applyBorder = applyBorder ? 1 : 0;
        if (applyNumberFormat !== undefined) xAttrs.applyNumberFormat = applyNumberFormat ? 1 : 0;
        if (applyAlignment !== undefined) xAttrs.applyAlignment = applyAlignment ? 1 : 0;
        if (xf.quotePrefix !== undefined) xAttrs.quotePrefix = xf.quotePrefix ? 1 : 0;
        if (xf.pivotButton !== undefined) xAttrs.pivotButton = xf.pivotButton ? 1 : 0;
        if (applyProtection !== undefined) xAttrs.applyProtection = applyProtection ? 1 : 0;

        const alignStr = xf.alignment ? this.alignmentXmlStr(xf.alignment) : "";
        const protStr = xf.protection ? this.protectionXmlStr(xf.protection) : "";
        const inner = alignStr + protStr;
        p.push(inner ? `<xf${attrs(xAttrs)}>${inner}</xf>` : `<xf${attrs(xAttrs)}/>`);
      }
      p.push("</cellXfs>");
    }

    // cellStyles — undefined = fresh document (emit the implicit Normal);
    // [] = adopted empty list (emit the empty container verbatim)
    if (this.customCellStyles !== undefined) {
      if (this.customCellStyles.length === 0) {
        p.push('<cellStyles count="0"/>');
      } else {
        // Normal (builtinId=0) is the implicit default; only auto-add it when
        // the caller's list doesn't already include it, so a parsed list
        // containing Normal round-trips instead of producing a duplicate.
        // Some sources omit builtinId — fall back to the name.
        const hasNormal = this.customCellStyles.some(
          (cs) => cs.builtinId === 0 || cs.name === "Normal",
        );
        const csAttrs: string[] = [`count="${this.customCellStyles.length + (hasNormal ? 0 : 1)}"`];
        const csParts: string[] = [`<cellStyles ${csAttrs.join(" ")}>`];
        if (!hasNormal) csParts.push('<cellStyle name="Normal" xfId="0" builtinId="0"/>');
        for (const cs of this.customCellStyles) {
          const attrs: string[] = [`name="${escapeXml(cs.name)}"`, `xfId="${cs.xfId}"`];
          if (cs.builtinId !== undefined) attrs.push(`builtinId="${cs.builtinId}"`);
          if (cs.customBuiltin !== undefined)
            attrs.push(`customBuiltin="${cs.customBuiltin ? 1 : 0}"`);
          if (cs.iLevel !== undefined) attrs.push(`iLevel="${cs.iLevel}"`);
          if (cs.hidden !== undefined) attrs.push(`hidden="${cs.hidden ? 1 : 0}"`);
          if (cs.uid !== undefined) attrs.push(`xr:uid="${escapeXml(cs.uid)}"`);
          csParts.push(`<cellStyle ${attrs.join(" ")}/>`);
        }
        csParts.push("</cellStyles>");
        p.push(csParts.join(""));
      }
    } else if (!this.roundTrip) {
      p.push(
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
      );
    }

    // dxfs
    if (this.dxfs.length > 0) {
      p.push(`<dxfs count="${this.dxfs.length}">`);
      for (const dxf of this.dxfs) {
        const dParts: string[] = [];
        // CT_Dxf sequence: font, numFmt, fill, alignment, border, protection.
        if (dxf.font) dParts.push(`<font>${this.fontXmlStr(dxf.font)}</font>`);
        if (dxf.numFmt) {
          const nf = typeof dxf.numFmt === "string" ? { formatCode: dxf.numFmt } : dxf.numFmt;
          // CT_NumFmt requires numFmtId; resolve from the built-in table and
          // fall back to the custom range when the code is not built-in.
          const numFmtId = nf.numFmtId ?? BUILTIN_NUMFMTS[nf.formatCode] ?? 164;
          dParts.push(`<numFmt numFmtId="${numFmtId}" formatCode="${escapeXml(nf.formatCode)}"/>`);
        }
        if (dxf.fill) {
          const f = dxf.fill;
          // Excel's dxf fill convention: the visible tint color carries on
          // bgColor (often with no patternType). Round-trips keep both color
          // elements as written; a fresh color-only fill (no bg channel) falls
          // back to bgColor so conditional formatting keeps showing the color.
          const patAttrs = f.patternType !== undefined ? attrs({ patternType: f.patternType }) : "";
          const hasBg =
            f.bgColor !== undefined ||
            f.bgColorIndexed !== undefined ||
            f.bgThemeColor !== undefined ||
            f.bgAutoColor !== undefined;
          // With a bg channel present the color field is the parsed fgColor;
          // alone it is the fresh-authoring color that lands on bgColor.
          const fgChannel =
            f.themeColor !== undefined
              ? `theme="${f.themeColor}"`
              : f.colorIndexed !== undefined
                ? `indexed="${f.colorIndexed}"`
                : f.color && (hasBg || f.fgColorRaw !== undefined)
                  ? `rgb="${f.fgColorRaw ?? `FF${f.color}`}"`
                  : f.fgAutoColor
                    ? 'auto="1"'
                    : "";
          const fgTint =
            f.tintRaw !== undefined
              ? ` tint="${f.tintRaw}"`
              : f.tint !== undefined
                ? ` tint="${decimalAttr(f.tint)}"`
                : "";
          const fgContent = fgChannel ? `<fgColor ${fgChannel}${fgTint}/>` : "";
          const bgChannel =
            f.bgThemeColor !== undefined
              ? `theme="${f.bgThemeColor}"`
              : f.bgColorIndexed !== undefined
                ? `indexed="${f.bgColorIndexed}"`
                : f.bgColor
                  ? `rgb="${f.bgColorRaw ?? `FF${f.bgColor}`}"`
                  : f.bgAutoColor
                    ? 'auto="1"'
                    : f.color && !hasBg && f.fgColorRaw === undefined
                      ? `rgb="${f.fgColorRaw ?? `FF${f.color}`}"`
                      : "";
          const bgTint =
            f.bgTintRaw !== undefined
              ? ` tint="${f.bgTintRaw}"`
              : f.bgTint !== undefined
                ? ` tint="${decimalAttr(f.bgTint)}"`
                : "";
          const bgContent = bgChannel ? `<bgColor ${bgChannel}${bgTint}/>` : "";
          const fillContent = fgContent + bgContent;
          dParts.push(
            fillContent
              ? `<fill><patternFill${patAttrs}>${fillContent}</patternFill></fill>`
              : `<fill><patternFill${patAttrs}/></fill>`,
          );
        }
        if (dxf.alignment) dParts.push(this.alignmentXmlStr(dxf.alignment));
        // dxf borders carry only the sides the source wrote — no padding
        if (dxf.border) {
          const borderAttrs = [
            ...(dxf.border.diagonalUp !== undefined
              ? [`diagonalUp="${dxf.border.diagonalUp ? 1 : 0}"`]
              : []),
            ...(dxf.border.diagonalDown !== undefined
              ? [`diagonalDown="${dxf.border.diagonalDown ? 1 : 0}"`]
              : []),
            ...(dxf.border.outline !== undefined
              ? [`outline="${dxf.border.outline ? 1 : 0}"`]
              : []),
          ].join(" ");
          dParts.push(
            `<border${borderAttrs ? ` ${borderAttrs}` : ""}>${this.borderXmlStr(dxf.border, false)}</border>`,
          );
        }
        if (dxf.protection) dParts.push(this.protectionXmlStr(dxf.protection));
        if (dxf.extensions && dxf.extensions.length > 0) {
          const extParts = dxf.extensions.map((ext) => {
            const ns = Object.entries(ext.namespaces ?? {})
              .map(([name, value]) => ` ${name}="${escapeXml(value)}"`)
              .join("");
            return ext.content
              ? `<ext uri="${escapeXml(ext.uri)}"${ns}>${ext.content}</ext>`
              : `<ext uri="${escapeXml(ext.uri)}"${ns}/>`;
          });
          dParts.push(`<extLst>${extParts.join("")}</extLst>`);
        }
        if (dParts.length > 0) {
          p.push(`<dxf>${dParts.join("")}</dxf>`);
        } else {
          p.push("<dxf/>");
        }
      }
      p.push("</dxfs>");
    } else if (this.dxfsDeclared || !this.roundTrip) {
      // Excel's fresh-file convention (and sources declaring an empty
      // container) write it; a round-trip source that omitted the section
      // keeps it omitted.
      p.push('<dxfs count="0"/>');
    }
    // tableStyles (CT_TableStyles)
    if (this.tableStyles) {
      const count = this.tableStyles.count ?? this.tableStyles.tableStyles?.length ?? 0;
      const defaultTableStyle = this.tableStyles.defaultTableStyle;
      const defaultPivotStyle = this.tableStyles.defaultPivotStyle;
      const tsParts: string[] = [
        `<tableStyles count="${count}"${defaultTableStyle !== undefined ? ` defaultTableStyle="${escapeXml(defaultTableStyle)}"` : ""}${defaultPivotStyle !== undefined ? ` defaultPivotStyle="${escapeXml(defaultPivotStyle)}"` : ""}>`,
      ];
      for (const ts of this.tableStyles.tableStyles ?? []) {
        const tsAttrs: string[] = [`name="${escapeXml(ts.name)}"`];
        if (ts.uid !== undefined) tsAttrs.push(`xr9:uid="${escapeXml(ts.uid)}"`);
        if (ts.pivot !== undefined) tsAttrs.push(`pivot="${ts.pivot ? 1 : 0}"`);
        if (ts.table !== undefined) tsAttrs.push(`table="${ts.table ? 1 : 0}"`);
        if (ts.elements && ts.elements.length > 0) {
          tsParts.push(`<tableStyle ${tsAttrs.join(" ")} count="${ts.elements.length}">`);
          for (const el of ts.elements) {
            const elAttrs: string[] = [`type="${el.type}"`];
            if (el.dxfId !== undefined) elAttrs.push(`dxfId="${el.dxfId}"`);
            if (el.size !== undefined) elAttrs.push(`size="${el.size}"`);
            tsParts.push(`<tableStyleElement ${elAttrs.join(" ")}/>`);
          }
          tsParts.push("</tableStyle>");
        } else {
          tsParts.push(`<tableStyle ${tsAttrs.join(" ")}/>`);
        }
      }
      tsParts.push("</tableStyles>");
      p.push(tsParts.join(""));
    } else if (this.tableStyles !== undefined || !this.roundTrip) {
      // Fresh files (and sources declaring an empty container) write the
      // defaults; a round-trip source without the section omits it.
      p.push(
        '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>',
      );
    }

    // colors (optional color palette)
    if (this.colors) {
      const c = this.colors;
      const colorParts: string[] = ["<colors>"];
      if (c.themeColors && c.themeColors.length > 0) {
        colorParts.push("<themeColors>");
        for (const value of c.themeColors) colorParts.push(`<rgbColor val="${value}"/>`);
        colorParts.push("</themeColors>");
      }
      if (c.indexedColors && c.indexedColors.length > 0) {
        colorParts.push("<indexedColors>");
        for (const ic of c.indexedColors) {
          colorParts.push(`<rgbColor rgb="${ic.rgb}"/>`);
        }
        colorParts.push("</indexedColors>");
      }
      if (c.legacyMruColors && c.legacyMruColors.length > 0) {
        colorParts.push("<mruColors>");
        for (const mc of c.legacyMruColors) {
          colorParts.push(`<color type="${mc.type}" val="${mc.value}"/>`);
        }
        colorParts.push("</mruColors>");
      } else if (c.mruColors && c.mruColors.length > 0) {
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
        const ns = Object.entries(ext.namespaces ?? {})
          .map(([name, value]) => ` ${name}="${value}"`)
          .join("");
        if (ext.content) {
          extParts.push(`<ext uri="${ext.uri}"${ns}>${ext.content}</ext>`);
        } else {
          extParts.push(`<ext uri="${ext.uri}"${ns}/>`);
        }
      }
      extParts.push("</extLst>");
      p.push(extParts.join(""));
    }

    p.push("</styleSheet>");
    return p.join("");
  }

  private fontXmlStr(f: FontOptions): string {
    const parts: { name: string; xml: string }[] = [];
    const push = (name: string, xml: string): void => {
      parts.push({ name, xml });
    };
    // CT_BooleanProperty val defaults to true — an explicit false round-trips
    // as val="0", an absent flag emits nothing.
    const flag = (name: string, on: boolean | undefined, raw?: string): void => {
      if (on === undefined) return;
      push(
        name,
        raw !== undefined ? `<${name} val="${raw}"/>` : on ? `<${name}/>` : `<${name} val="0"/>`,
      );
    };
    // x:font child order follows Excel's writer (and the Open XML SDK particle):
    // b, i, strike, condense, extend, outline, shadow, u, vertAlign, sz, color,
    // name, family, charset, scheme — the literal ISO sequence (name/charset/
    // family leading) matches no real-world file
    flag("b", f.bold, f.boldRaw);
    flag("i", f.italic, f.italicRaw);
    flag("strike", f.strike, f.strikeRaw);
    flag("condense", f.condense, f.condenseRaw);
    flag("extend", f.extend, f.extendRaw);
    flag("outline", f.outline, f.outlineRaw);
    flag("shadow", f.shadow, f.shadowRaw);
    if (f.underline !== undefined) {
      push(
        "u",
        f.underline === false
          ? '<u val="none"/>'
          : f.underline === true
            ? "<u/>"
            : `<u val="${f.underline}"/>`,
      );
    }
    if (f.vertAlign) push("vertAlign", `<vertAlign val="${f.vertAlign}"/>`);
    if (f.size !== undefined) push("sz", `<sz val="${f.sizeRaw ?? f.size}"/>`);
    if (f.autoColor) push("color", '<color auto="1"/>');
    else if (f.legacyColorType !== undefined)
      push(
        "color",
        `<color type="${f.legacyColorType}" val="${f.legacyColorValue ?? ""}"${
          f.tintRaw !== undefined
            ? ` tint="${f.tintRaw}"`
            : f.tint !== undefined
              ? ` tint="${decimalAttr(f.tint)}"`
              : ""
        }/>`,
      );
    else if (f.themeColor !== undefined)
      push(
        "color",
        `<color theme="${f.themeColor}"${
          f.tintRaw !== undefined
            ? ` tint="${f.tintRaw}"`
            : f.tint !== undefined
              ? ` tint="${decimalAttr(f.tint)}"`
              : ""
        }/>`,
      );
    else if (f.colorIndexed !== undefined) push("color", `<color indexed="${f.colorIndexed}"/>`);
    else if (f.color)
      push(
        "color",
        `<color rgb="${f.colorRaw ?? `FF${f.color}`}"${
          f.tintRaw !== undefined
            ? ` tint="${f.tintRaw}"`
            : f.tint !== undefined
              ? ` tint="${decimalAttr(f.tint)}"`
              : ""
        }/>`,
      );
    if (f.font) push("name", `<name val="${escapeXml(f.font)}"/>`);
    if (f.family !== undefined) push("family", `<family val="${f.family}"/>`);
    if (f.charset !== undefined) push("charset", `<charset val="${f.charset}"/>`);
    if (f.scheme) push("scheme", `<scheme val="${f.scheme}"/>`);
    const order = f.childOrder;
    if (order !== undefined) {
      const rank = new Map(order.map((name, index) => [name, index] as const));
      parts.sort(
        (left, right) =>
          (rank.get(left.name) ?? order.length) - (rank.get(right.name) ?? order.length),
      );
    }
    return parts.map((part) => part.xml).join("");
  }

  /**
   * Serialize a border. `allSides` pads the five cell sides Excel always
   * writes (cellXfs context); dxf borders carry only the sides present.
   */
  private borderXmlStr(b: BorderSideOptions, allSides = true): string {
    const parts: { name: string; xml: string }[] = [];
    const sideColorXmlStr = (side: BorderOptions): string => {
      if (side.autoColor) return '<color auto="1"/>';
      if (side.legacyColorType !== undefined)
        return `<color type="${side.legacyColorType}" val="${side.legacyColorValue ?? ""}"${
          side.tintRaw !== undefined
            ? ` tint="${side.tintRaw}"`
            : side.tint !== undefined
              ? ` tint="${decimalAttr(side.tint)}"`
              : ""
        }/>`;
      if (side.themeColor !== undefined)
        return `<color theme="${side.themeColor}"${
          side.tintRaw !== undefined
            ? ` tint="${side.tintRaw}"`
            : side.tint !== undefined
              ? ` tint="${decimalAttr(side.tint)}"`
              : ""
        }/>`;
      if (side.colorIndexed !== undefined) return `<color indexed="${side.colorIndexed}"/>`;
      if (side.color) return `<color rgb="${side.colorRaw ?? `FF${side.color}`}"/>`;
      return "";
    };
    const renderSide = (name: string, opts: BorderOptions | undefined, required: boolean) => {
      const sideContent = opts ? sideColorXmlStr(opts) : "";
      if (opts?.style !== undefined) {
        parts.push({ name, xml: `<${name} style="${opts.style}">${sideContent}</${name}>` });
      } else if (opts) {
        parts.push({ name, xml: `<${name}>${sideContent}</${name}>` });
      } else if (required) {
        parts.push({ name, xml: `<${name}/>` });
      }
    };
    const anySidePresent = [
      b.left,
      b.right,
      b.top,
      b.bottom,
      b.diagonal,
      b.vertical,
      b.horizontal,
      b.start,
      b.end,
    ].some((side) => side !== undefined);
    const cellSides = ["left", "right", "top", "bottom", "diagonal"] as const;
    for (const side of cellSides) {
      renderSide(side, b[side] as BorderOptions | undefined, allSides && !anySidePresent);
    }
    // vertical/horizontal/start/end: cell-range or RTL sides, only when present
    renderSide("vertical", b.vertical, false);
    renderSide("horizontal", b.horizontal, false);
    // start/end not in transitional XSD — only emit when present
    renderSide("start", b.start, false);
    renderSide("end", b.end, false);
    const order = b.childOrder;
    if (order !== undefined) {
      const rank = new Map(order.map((name, index) => [name, index] as const));
      parts.sort(
        (left, right) =>
          (rank.get(left.name) ?? order.length) - (rank.get(right.name) ?? order.length),
      );
    }
    return parts.map((part) => part.xml).join("");
  }

  private alignmentXmlStr(a: AlignmentOptions): string {
    const aAttrs: Record<string, string | number | boolean | undefined> = {};
    if (a.horizontal) aAttrs.horizontal = a.horizontal;
    if (a.vertical) aAttrs.vertical = a.vertical;
    if (a.wrapText !== undefined) aAttrs.wrapText = a.wrapText ? 1 : 0;
    if (a.textRotation !== undefined) aAttrs.textRotation = a.textRotation;
    if (a.indent !== undefined) aAttrs.indent = a.indent;
    if (a.indentRaw !== undefined) aAttrs.indent = a.indentRaw;
    if (a.relativeIndent !== undefined) aAttrs.relativeIndent = a.relativeIndent;
    if (a.justifyLastLine !== undefined) aAttrs.justifyLastLine = a.justifyLastLine ? 1 : 0;
    if (a.shrinkToFit !== undefined) aAttrs.shrinkToFit = a.shrinkToFit ? 1 : 0;
    if (a.readingOrder !== undefined) aAttrs.readingOrder = a.readingOrder;
    if (a.mergeCell !== undefined) aAttrs.mergeCell = a.mergeCell ? 1 : 0;
    return `<alignment${attrs(aAttrs)}/>`;
  }

  private protectionXmlStr(pr: CellProtectionOptions): string {
    const prAttrs: Record<string, string | number | boolean | undefined> = {};
    if (pr.locked !== undefined) prAttrs.locked = pr.locked ? 1 : 0;
    if (pr.hidden !== undefined) prAttrs.hidden = pr.hidden ? 1 : 0;
    return `<protection${attrs(prAttrs)}/>`;
  }
}
