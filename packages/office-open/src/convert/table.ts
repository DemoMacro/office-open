/**
 * Cross-format table conversion.
 *
 * docx (w:tbl flow) and pptx (a:tbl graphic) share a structural core
 * (rows/cells/span/6-flags/columnWidths/vertical-align) defined in
 * `@office-open/core`'s BaseTableOptions; both packages extend it. The
 * structural fields pass through directly and only the per-package parts need
 * translation:
 * - cell content: docx w:p (SectionChild) ↔ pptx a:p (ParagraphDescriptor),
 *   via ./text. docx cell non-paragraph children (nested table/toc/sdt/…)
 *   drop (flatten not implemented).
 * - column widths: `number` is the native unit (docx twip, pptx EMU) so it is
 *   converted (×635 / ÷635); UniversalMeasure strings pass through (each
 *   package's descriptor resolves them).
 * - position: pptx absolute x/y; docx is flow (no position) — lost pptx→docx.
 * - styles (fill/borders/margins): w:/a: domain types differ; only solid fill
 *   maps (pptx→docx), the rest is dropped (matches MS Office paste loss).
 * - row height: docx {value,rule} twip ↔ pptx number EMU.
 *
 * xlsx has no visual table object (its sml Table is a data range), so docx/pptx
 * tables restore to worksheet fragments: cell value = first-paragraph plain
 * text, mergeCells = columnSpan/rowSpan, column widths / row heights converted
 * to xlsx units (character width / points). Cell styles are dropped (dxf
 * synthesis is a follow-up). The reverse takes xlsx fragments back to a table.
 *
 * @module
 */
import {
  convertEmuToPoints,
  convertEmuToTwip,
  convertPointsToEmu,
  convertPointsToTwip,
  convertToEmu,
  convertToPt,
  convertToTwip,
  convertTwipToEmu,
  ThemeColor,
} from "@office-open/core";
import type { UniversalMeasure } from "@office-open/core";
import type { ParagraphDescriptorOptions } from "@office-open/core/drawing";
import type {
  ParagraphOptions,
  SectionChild,
  TableOptions as DocxTableOptions,
  TableCellOptions as DocxTableCellOptions,
  TableRowOptions as DocxTableRowOptions,
} from "@office-open/docx";
import type {
  TableOptions as PptxTableOptions,
  TableCellOptions as PptxTableCellOptions,
  TableRowOptions as PptxTableRowOptions,
} from "@office-open/pptx";
import type {
  CellOptions as XlsxCellOptions,
  ColumnOptions,
  MergeCellOptions,
  RowOptions as XlsxRowOptions,
} from "@office-open/xlsx";
import { columnToLetter, letterToColumn } from "@office-open/xlsx";

import { DEFAULT_COL_EMU } from "./position";
import { fromDrawingParagraph, toDrawingParagraph } from "./text";

/** Heuristic EMU per character of column width (8.43 chars ≈ DEFAULT_COL_EMU). */
const EMU_PER_CHAR = DEFAULT_COL_EMU / 8.43;

/** xlsx visual-table restoration: worksheet fragments (rows/merges/columns). */
export interface XlsxVisualTable {
  rows: XlsxRowOptions[];
  mergeCells?: MergeCellOptions[];
  columns?: ColumnOptions[];
}

/** Parse a merge ref ("A1:D1") into 0-based row/col corners. */
function parseMergeRef(
  ref: string,
): { row: number; col: number; rowEnd: number; colEnd: number } | undefined {
  const [from = "", to = from] = ref.split(":");
  const fa = from.match(/^([A-Z]+)(\d+)$/);
  const fb = to.match(/^([A-Z]+)(\d+)$/);
  if (!fa || !fb) return undefined;
  return {
    row: Number(fa[2]) - 1,
    col: letterToColumn(fa[1]) - 1,
    rowEnd: Number(fb[2]) - 1,
    colEnd: letterToColumn(fb[1]) - 1,
  };
}

// ── unit helpers ──
// EMU↔twip↔points conversions live in @office-open/core; only the xlsx
// character-width heuristic (EMU_PER_CHAR) is local to visual restoration.

const emuToCharWidth = (emu: number): number => emu / EMU_PER_CHAR;
const charWidthToEmu = (chars: number): number => chars * EMU_PER_CHAR;

/** docx height value (twip number or UM) → EMU. */
const docxHeightToEmu = (v: number | UniversalMeasure): number =>
  typeof v === "number" ? convertTwipToEmu(v) : convertToEmu(v);
/** pptx height value (EMU number or UM) → twip. */
const pptxHeightToTwip = (v: number | UniversalMeasure): number =>
  typeof v === "number" ? convertEmuToTwip(v) : convertToTwip(v);

/** Convert numeric column widths via `convert`; UM strings pass through. */
function convertColumnWidths(
  widths: (number | string)[] | undefined,
  convert: (n: number) => number,
): (number | string)[] | undefined {
  if (!widths) return undefined;
  return widths.map((w) => (typeof w === "number" ? convert(w) : w));
}

// ── discriminant ──

/** Structural test: an XlsxVisualTable's first cell lacks docx/pptx markers
 *  (children/text/shading), or it carries xlsx-only columns/mergeCells. */
function isXlsxVisual(src: unknown): src is XlsxVisualTable {
  if (typeof src !== "object" || src === null) return false;
  const s = src as Record<string, unknown>;
  if (Array.isArray(s.columns) || Array.isArray(s.mergeCells)) return true;
  const firstCell = (s.rows as { cells?: Array<Record<string, unknown>> }[] | undefined)?.[0]
    ?.cells?.[0];
  if (!firstCell) return false;
  return !("children" in firstCell) && !("text" in firstCell) && !("shading" in firstCell);
}

/** Copy the 6 special-row flags (matching field names across docx/pptx). */
function copyBaseTableFlags<S extends DocxTableOptions | PptxTableOptions>(
  src: S,
): Pick<
  DocxTableOptions & PptxTableOptions,
  "firstRow" | "lastRow" | "firstCol" | "lastCol" | "bandRow" | "bandCol"
> {
  return {
    ...(src.firstRow !== undefined ? { firstRow: src.firstRow } : {}),
    ...(src.lastRow !== undefined ? { lastRow: src.lastRow } : {}),
    ...(src.firstCol !== undefined ? { firstCol: src.firstCol } : {}),
    ...(src.lastCol !== undefined ? { lastCol: src.lastCol } : {}),
    ...(src.bandRow !== undefined ? { bandRow: src.bandRow } : {}),
    ...(src.bandCol !== undefined ? { bandCol: src.bandCol } : {}),
  };
}

// ── cell content bridges ──

/** docx cell (SectionChild[], w:p) → pptx paragraphs (a:p). */
function docxCellToPptxContent(
  cell: DocxTableCellOptions,
): (ParagraphDescriptorOptions | string)[] {
  const out: (ParagraphDescriptorOptions | string)[] = [];
  for (const child of cell.children ?? []) {
    if (typeof child === "string") {
      out.push(child);
    } else if ("paragraph" in child) {
      const para: ParagraphOptions =
        typeof child.paragraph === "string"
          ? { children: [{ text: child.paragraph }] }
          : child.paragraph;
      out.push(toDrawingParagraph(para));
    }
    // nested table/toc/sdt/… → drop
  }
  return out;
}

/** pptx cell (a:p) → docx SectionChild[] (w:p). */
function pptxCellToDocxChildren(cell: PptxTableCellOptions): SectionChild[] {
  const out: SectionChild[] = [];
  if (cell.text !== undefined) {
    out.push({ paragraph: { children: [{ text: cell.text }] } });
  }
  for (const child of cell.children ?? []) {
    if (typeof child === "string") {
      out.push({ paragraph: { children: [{ text: child }] } });
    } else {
      out.push({ paragraph: fromDrawingParagraph(child) });
    }
  }
  return out;
}

/** First-paragraph plain text from a docx cell. */
function docxCellText(cell: DocxTableCellOptions): string | undefined {
  for (const child of cell.children ?? []) {
    if (typeof child === "string") return child;
    if ("paragraph" in child) {
      const para =
        typeof child.paragraph === "string"
          ? { children: [{ text: child.paragraph }] }
          : child.paragraph;
      const text = para.children
        ?.map((r) => (typeof r === "string" ? r : "text" in r ? (r.text ?? "") : ""))
        .join("");
      if (text) return text;
    }
  }
  return undefined;
}

/** Plain text from a pptx cell. */
function pptxCellText(cell: PptxTableCellOptions): string | undefined {
  if (cell.text !== undefined) return cell.text;
  for (const child of cell.children ?? []) {
    if (typeof child === "string") return child;
    const text = child.children
      ?.map((r) => (typeof r === "string" ? r : "text" in r ? (r.text ?? "") : ""))
      .join("");
    if (text) return text;
  }
  return undefined;
}

/** Map a: scheme color token (ST_SchemeColorVal) → w: themeColor token
 *  (ST_ThemeColor). accent1-6 pass through; bg/tx/dk/lt → background/text/
 *  dark/light; hlink → hyperlink, folHlink → followedHyperlink. phClr has no
 *  w: equivalent (dropped). */
const SCHEME_TO_THEME: Record<string, (typeof ThemeColor)[keyof typeof ThemeColor]> = {
  bg1: "background1",
  tx1: "text1",
  bg2: "background2",
  tx2: "text2",
  dk1: "dark1",
  lt1: "light1",
  dk2: "dark2",
  lt2: "light2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hlink: "hyperlink",
  folHlink: "followedHyperlink",
};

/** pptx FillOptions (a:fill) → docx ShadingProperties (w:shd). RGB hex and
 *  RgbColorOptions → `@fill`; SchemeColorOptions → `@themeColor` (token mapped);
 *  hsl/system/preset/scRgb/phClr and color transforms → dropped (docx shading
 *  is RGB-hex or theme-color only). */
function pptxFillToDocxShading(
  fill: PptxTableCellOptions["fill"],
): DocxTableCellOptions["shading"] | undefined {
  if (fill === undefined) return undefined;
  if (typeof fill === "string") return { fill };
  if (fill.type === "solid") {
    if (typeof fill.color === "string") return { fill: fill.color };
    if (typeof fill.color === "object" && "value" in fill.color) {
      const v = fill.color.value;
      if (typeof v === "string") {
        return v in SCHEME_TO_THEME ? { themeColor: SCHEME_TO_THEME[v] } : { fill: v };
      }
    }
    return undefined; // hsl/system/preset/scRgb → drop
  }
  return undefined; // gradient/pattern/blip → drop
}

/** xlsx cell value → plain text. Primitives stringify directly; Date → ISO;
 *  RichTextOptions falls back to JSON (rich-text run extraction is a follow-up). */
function cellValueToText(value: XlsxCellOptions["value"]): string | undefined {
  if (value === null || value === undefined) return undefined;
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  if (value instanceof Date) return value.toISOString();
  return JSON.stringify(value);
}

// ── → docx ──

/** Convert a pptx table to a docx table. */
export function toDocxTable(source: PptxTableOptions): DocxTableOptions;
/** Convert xlsx visual-table fragments back to a docx table. */
export function toDocxTable(source: XlsxVisualTable): DocxTableOptions;
export function toDocxTable(source: PptxTableOptions | XlsxVisualTable): DocxTableOptions {
  return isXlsxVisual(source) ? xlsxToDocx(source) : pptxToDocx(source);
}

function pptxToDocx(src: PptxTableOptions): DocxTableOptions {
  const rows: DocxTableRowOptions[] = src.rows.map((row) => ({
    ...(row.height !== undefined ? { height: { value: pptxHeightToTwip(row.height) } } : {}),
    cells: row.cells.map((cell): DocxTableCellOptions => {
      const shading = pptxFillToDocxShading(cell.fill);
      return {
        children: pptxCellToDocxChildren(cell),
        ...(cell.columnSpan !== undefined ? { columnSpan: cell.columnSpan } : {}),
        ...(cell.rowSpan !== undefined ? { rowSpan: cell.rowSpan } : {}),
        ...(cell.verticalAlign !== undefined
          ? {
              verticalAlign:
                cell.verticalAlign === "justify" || cell.verticalAlign === "distribute"
                  ? "center"
                  : cell.verticalAlign,
            }
          : {}),
        ...(shading ? { shading } : {}),
      };
    }),
  }));
  return {
    rows,
    ...(src.columnWidths
      ? { columnWidths: convertColumnWidths(src.columnWidths, convertEmuToTwip) as number[] }
      : {}),
    ...copyBaseTableFlags(src),
  };
}

function xlsxToDocx(src: XlsxVisualTable): DocxTableOptions {
  const merges = (src.mergeCells ?? [])
    .map((m) => parseMergeRef(m.ref))
    .filter((m): m is NonNullable<typeof m> => m !== undefined);
  const rows: DocxTableRowOptions[] = src.rows.map((row, ri) => ({
    ...(row.height !== undefined
      ? { height: { value: convertPointsToTwip(convertToPt(row.height)) } }
      : {}),
    cells: (row.cells ?? []).map((cell, ci): DocxTableCellOptions => {
      const merge = merges.find((m) => m.row === ri && m.col === ci);
      const columnSpan = merge ? merge.colEnd - merge.col + 1 : undefined;
      const rowSpan = merge ? merge.rowEnd - merge.row + 1 : undefined;
      const text = cellValueToText(cell.value);
      return {
        children: text !== undefined ? [{ paragraph: { children: [{ text }] } }] : [],
        ...(columnSpan && columnSpan > 1 ? { columnSpan } : {}),
        ...(rowSpan && rowSpan > 1 ? { rowSpan } : {}),
      };
    }),
  }));
  return {
    rows,
    ...(src.columns
      ? {
          columnWidths: src.columns.map((c) => convertEmuToTwip(charWidthToEmu(c.width ?? 8.43))),
        }
      : {}),
  };
}

// ── → pptx ──

/** Convert a docx table to a pptx table. */
export function toPptxTable(source: DocxTableOptions): PptxTableOptions;
/** Convert xlsx visual-table fragments back to a pptx table. */
export function toPptxTable(source: XlsxVisualTable): PptxTableOptions;
export function toPptxTable(source: DocxTableOptions | XlsxVisualTable): PptxTableOptions {
  return isXlsxVisual(source) ? xlsxToPptx(source) : docxToPptx(source);
}

function docxToPptx(src: DocxTableOptions): PptxTableOptions {
  const rows: PptxTableRowOptions[] = src.rows.map((row) => {
    if (!("cells" in row)) return { cells: [] }; // sdt/customXml row → flatten
    return {
      ...(row.height ? { height: docxHeightToEmu(row.height.value) } : {}),
      cells: row.cells.map((cell): PptxTableCellOptions => {
        if (!("children" in cell)) return { children: [] }; // sdt/customXml cell → flatten
        return {
          children: docxCellToPptxContent(cell),
          ...(cell.columnSpan !== undefined ? { columnSpan: cell.columnSpan } : {}),
          ...(cell.rowSpan !== undefined ? { rowSpan: cell.rowSpan } : {}),
          ...(cell.verticalAlign !== undefined ? { verticalAlign: cell.verticalAlign } : {}),
        };
      }),
    };
  });
  return {
    rows,
    ...(src.columnWidths
      ? { columnWidths: convertColumnWidths(src.columnWidths, convertTwipToEmu) as number[] }
      : {}),
    ...copyBaseTableFlags(src),
  };
}

function xlsxToPptx(src: XlsxVisualTable): PptxTableOptions {
  const merges = (src.mergeCells ?? [])
    .map((m) => parseMergeRef(m.ref))
    .filter((m): m is NonNullable<typeof m> => m !== undefined);
  const rows: PptxTableRowOptions[] = src.rows.map((row, ri) => ({
    ...(row.height !== undefined ? { height: convertPointsToEmu(convertToPt(row.height)) } : {}),
    cells: (row.cells ?? []).map((cell, ci): PptxTableCellOptions => {
      const merge = merges.find((m) => m.row === ri && m.col === ci);
      const columnSpan = merge ? merge.colEnd - merge.col + 1 : undefined;
      const rowSpan = merge ? merge.rowEnd - merge.row + 1 : undefined;
      const text = cellValueToText(cell.value);
      return {
        ...(text !== undefined ? { text } : {}),
        ...(columnSpan && columnSpan > 1 ? { columnSpan } : {}),
        ...(rowSpan && rowSpan > 1 ? { rowSpan } : {}),
      };
    }),
  }));
  return {
    rows,
    ...(src.columns
      ? { columnWidths: src.columns.map((c) => Math.round(charWidthToEmu(c.width ?? 8.43))) }
      : {}),
  };
}

// ── → xlsx ──

/** Convert a docx table to xlsx worksheet fragments (visual restoration). */
export function toXlsxTable(source: DocxTableOptions): XlsxVisualTable;
/** Convert a pptx table to xlsx worksheet fragments (visual restoration). */
export function toXlsxTable(source: PptxTableOptions): XlsxVisualTable;
export function toXlsxTable(source: DocxTableOptions | PptxTableOptions): XlsxVisualTable {
  return isDocxTable(source) ? docxToXlsx(source) : pptxToXlsx(source);
}

/** docx vs pptx: both extend BaseTableOptions so top-level fields mostly overlap.
 *  Decide by content shape — docx wraps cell content as SectionChild
 *  ({ paragraph | table | toc | … }), pptx stores flat a:p paragraphs or a
 *  `text` shorthand — and by domain-only keys (docx w:shd/float/style; pptx
 *  a:fill/tableStyleId). `width` overlaps (docx TableWidthProperties vs pptx
 *  number) so it is intentionally not used. */
function isDocxTable(src: unknown): src is DocxTableOptions {
  if (typeof src !== "object" || src === null) return false;
  const s = src as Record<string, unknown>;
  if ("tableStyleId" in s) return false; // pptx-only
  if ("style" in s || "float" in s || "visuallyRightToLeft" in s || "indent" in s) return true; // docx-only
  const rows = s.rows as Array<Record<string, unknown>> | undefined;
  if (!Array.isArray(rows)) return false;
  for (const row of rows) {
    if (!("cells" in row) || !Array.isArray(row.cells)) continue;
    for (const cell of row.cells as Array<Record<string, unknown>>) {
      if ("shading" in cell) return true; // docx w:shd
      if ("text" in cell || "fill" in cell) return false; // pptx a:fill / shorthand
      const child = (cell.children as Array<Record<string, unknown>> | undefined)?.[0];
      if (child && typeof child === "object") {
        return (
          "paragraph" in child ||
          "table" in child ||
          "toc" in child ||
          "sdt" in child ||
          "customXml" in child ||
          "altChunk" in child
        );
      }
    }
  }
  return false;
}

function docxToXlsx(src: DocxTableOptions): XlsxVisualTable {
  const mergeCells: MergeCellOptions[] = [];
  const rows: XlsxRowOptions[] = src.rows.map((row, ri) => {
    if (!("cells" in row)) return { cells: [] };
    let ci = 0;
    const cells: XlsxCellOptions[] = row.cells.map((cell): XlsxCellOptions => {
      if (!("children" in cell)) {
        ci += 1;
        return {};
      }
      const span = cell.columnSpan ?? 1;
      const rspan = cell.rowSpan ?? 1;
      if (span > 1 || rspan > 1) {
        mergeCells.push({
          ref: `${columnToLetter(ci + 1)}${ri + 1}:${columnToLetter(ci + span)}${ri + rspan}`,
        });
      }
      const value = docxCellText(cell);
      ci += span;
      return value !== undefined ? { value } : {};
    });
    const xrow: XlsxRowOptions = { cells };
    if (row.height) xrow.height = convertEmuToPoints(docxHeightToEmu(row.height.value));
    return xrow;
  });
  const columns: ColumnOptions[] | undefined = src.columnWidths
    ? src.columnWidths.map((w, i) => ({
        min: i + 1,
        max: i + 1,
        width: typeof w === "number" ? emuToCharWidth(convertTwipToEmu(w)) : undefined,
      }))
    : undefined;
  return { rows, ...(mergeCells.length ? { mergeCells } : {}), ...(columns ? { columns } : {}) };
}

function pptxToXlsx(src: PptxTableOptions): XlsxVisualTable {
  const mergeCells: MergeCellOptions[] = [];
  const rows: XlsxRowOptions[] = src.rows.map((row, ri) => {
    let ci = 0;
    const cells: XlsxCellOptions[] = row.cells.map((cell): XlsxCellOptions => {
      const span = cell.columnSpan ?? 1;
      const rspan = cell.rowSpan ?? 1;
      if (span > 1 || rspan > 1) {
        mergeCells.push({
          ref: `${columnToLetter(ci + 1)}${ri + 1}:${columnToLetter(ci + span)}${ri + rspan}`,
        });
      }
      const value = pptxCellText(cell);
      ci += span;
      return value !== undefined ? { value } : {};
    });
    const xrow: XlsxRowOptions = { cells };
    if (row.height !== undefined)
      xrow.height = convertEmuToPoints(
        typeof row.height === "number" ? row.height : convertToEmu(row.height),
      );
    return xrow;
  });
  const columns: ColumnOptions[] | undefined = src.columnWidths
    ? src.columnWidths.map((w, i) => ({
        min: i + 1,
        max: i + 1,
        width: typeof w === "number" ? emuToCharWidth(w) : undefined,
      }))
    : undefined;
  return { rows, ...(mergeCells.length ? { mergeCells } : {}), ...(columns ? { columns } : {}) };
}
