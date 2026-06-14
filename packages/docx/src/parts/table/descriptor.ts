/**
 * Table (w:tbl) descriptor for DOCX.
 *
 * Stringifies pure JSON TableOptions into XML using direct string
 * concatenation — zero IXmlableObject, zero xml(), zero BaseXmlComponent.
 *
 * @module
 */

import { ThemeColor } from "@office-open/core";
import { xsdVerticalMergeRev } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrBool, attrNum, children, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { stringifySdtShell } from "@parts/bodychildren";
import { stringifyParagraphInline } from "@parts/inline";
import { parseRunProperties } from "@parts/paragraph/run/run-parse";
import { parseSdtProperties } from "@parts/sdt/sdt-parse";
import type { TableGridChangeOptions } from "@parts/table/grid";
import type { TableOptions } from "@parts/table/table";
import type { SdtCellOptions, TableCellOptions } from "@parts/table/table-cell/table-cell";
import { VerticalMergeType } from "@parts/table/table-cell/table-cell-components";
import type { TableBordersOptions } from "@parts/table/table-properties/table-borders";
import type { TableLookOptions } from "@parts/table/table-properties/table-look";
import type { SdtRowOptions, TableRowOptions } from "@parts/table/table-row/table-row";
import { BorderStyle } from "@shared/border";
import type { BorderOptions } from "@shared/border";
import type { SectionChild } from "@shared/section";

import type { BodyContext, DocxReadContext } from "../../context";
import {
  stringifyTableCellProperties,
  stringifyTableProperties,
  stringifyTablePropertyExceptions,
  stringifyTableRowProperties,
  type TablePropertiesOptions,
} from "./stringify";

// Valid border @w:val (ST_Border / BorderStyle) and @w:themeColor (ST_ThemeColor) values.
const BORDER_STYLES = Object.values(BorderStyle) as readonly string[];
const THEME_COLORS = Object.values(ThemeColor) as readonly string[];

/** Parse track-change attributes (id/author/date) from w:ins/w:del/w:cellIns/w:cellDel. */
function parseChangeAttrs(el: Element): Record<string, unknown> {
  const change: Record<string, unknown> = {};
  const id = attrNum(el, "w:id");
  if (id !== undefined) change.id = id;
  const author = attr(el, "w:author");
  if (author) change.author = author;
  const date = attr(el, "w:date");
  if (date) change.date = date;
  return change;
}

// ── Table grid ──

function buildTableGridXml(widths: number[], revision?: TableGridChangeOptions): string {
  const cols = widths.map((w) => `<w:gridCol w:w="${w}"/>`).join("");

  if (revision) {
    const revCols = revision.columnWidths.map((w) => `<w:gridCol w:w="${w}"/>`).join("");
    return `<w:tblGrid>${cols}<w:tblGridChange w:id="${revision.id}"><w:tblGrid>${revCols}</w:tblGrid></w:tblGridChange></w:tblGrid>`;
  }

  return `<w:tblGrid>${cols}</w:tblGrid>`;
}

// ── Cell span extraction ──

/** Extract column/row span from a plain-object cell. */
function getCellSpans(cell: TableCellOptions): {
  columnSpan: number;
  rowSpan: number;
} {
  return { columnSpan: cell.columnSpan ?? 1, rowSpan: cell.rowSpan ?? 1 };
}

// ── Cell children stringification ──

/**
 * Stringify a cell child (SectionChild) inline.
 * Handles strings and plain objects (paragraph/table).
 */
function stringifyCellChild(child: SectionChild, ctx: BodyContext): string {
  if (typeof child === "string") {
    return stringifyParagraphInline(child, ctx);
  }

  // Plain object dispatch
  if ("paragraph" in child) {
    return stringifyParagraphInline(child.paragraph, ctx);
  }
  if ("table" in child) {
    // Recursive: nested table via descriptor
    return tableDesc.stringify(child.table, ctx) ?? "";
  }

  // Fallback for other types — should not happen inside table cells
  return "";
}

// ── Cell stringification ──

function stringifyTableCell(cell: TableCellOptions, ctx: BodyContext): string {
  const parts: string[] = [];

  const tcPr = stringifyTableCellProperties(cell);
  if (tcPr) parts.push(tcPr);

  const children = cell.children as SectionChild[] | undefined;
  if (children) {
    for (const child of children) {
      parts.push(stringifyCellChild(child, ctx));
    }
  }

  // Cells must end with a paragraph unless the last child is already one
  const last = children?.[children.length - 1];
  const endsWithParagraph =
    last && typeof last !== "string" && ("paragraph" in last || "table" in last);
  if (!endsWithParagraph) {
    parts.push("<w:p/>");
  }

  return `<w:tc>${parts.join("")}</w:tc>`;
}

// ── Row stringification ──

function stringifyTableRow(
  row: TableRowOptions,
  ctx: BodyContext,
  extraCells?: { cell: TableCellOptions; columnIndex: number }[],
): string {
  const parts: string[] = [];

  // Property exceptions (tblPrEx)
  if (row.propertyExceptions) {
    parts.push(stringifyTablePropertyExceptions(row.propertyExceptions));
  }

  // Row properties
  const trPr = stringifyTableRowProperties(row);
  if (trPr) parts.push(trPr);

  const prefixCount = parts.length;

  // Cells (a cell may itself be wrapped by a cell-level SDT)
  for (const cell of row.cells) {
    if ("sdt" in cell) {
      const s = cell.sdt;
      const contentXml = (s.cells ?? []).map((c) => stringifyTableCell(c, ctx)).join("");
      parts.push(stringifySdtShell(s.properties, s.endProperties, contentXml));
    } else {
      parts.push(stringifyTableCell(cell, ctx));
    }
  }

  // Insert extra CONTINUE cells at correct positions
  if (extraCells && extraCells.length > 0) {
    for (const { cell, columnIndex } of extraCells) {
      const insertIdx = findInsertIndex(row.cells, columnIndex, prefixCount);
      parts.splice(insertIdx, 0, stringifyTableCell(cell, ctx));
    }
  }

  // rsid attributes
  const rsidAttrs: string[] = [];
  if (row.rsidRPr) rsidAttrs.push(` w:rsidRPr="${row.rsidRPr}"`);
  if (row.rsidR) rsidAttrs.push(` w:rsidR="${row.rsidR}"`);
  if (row.rsidDel) rsidAttrs.push(` w:rsidDel="${row.rsidDel}"`);
  if (row.rsidTr) rsidAttrs.push(` w:rsidTr="${row.rsidTr}"`);
  const attr = rsidAttrs.join("");

  const body = parts.join("");
  return body ? `<w:tr${attr}>${body}</w:tr>` : attr ? `<w:tr${attr}/>` : "<w:tr/>";
}

// ── Row options type ──

function findInsertIndex(
  cells: TableRowOptions["cells"],
  columnIndex: number,
  prefixCount: number,
): number {
  let colIdx = 0;
  for (let i = 0; i < cells.length; i++) {
    const c = cells[i];
    if ("sdt" in c) continue; // SDT-wrapped cells don't occupy grid columns here
    const { columnSpan } = getCellSpans(c);
    colIdx += columnSpan;
    if (colIdx > columnIndex) {
      return i + prefixCount;
    }
  }
  return cells.length + prefixCount;
}

// ── Vertical merge ──

/**
 * Pre-process rows to compute CONTINUE cells for vertical merge.
 */
function computeVerticalMergeCells(
  rows: (TableRowOptions | { sdt: SdtRowOptions })[],
): Map<number, { cell: TableCellOptions; columnIndex: number }[]> {
  const extraMap = new Map<number, { cell: TableCellOptions; columnIndex: number }[]>();
  for (let ri = 0; ri < rows.length - 1; ri++) {
    const row = rows[ri];
    if ("sdt" in row) continue; // SDT-wrapped rows don't participate in merge
    const cells = row.cells;
    let colIdx = 0;

    for (const cell of cells) {
      if ("sdt" in cell) continue; // SDT-wrapped cells don't participate in merge
      const typedCell = cell;
      const { columnSpan, rowSpan } = getCellSpans(typedCell);

      if (rowSpan > 1) {
        const continueCell: TableCellOptions = {
          borders: typedCell.borders,
          children: [],
          columnSpan,
          rowSpan: rowSpan - 1,
          verticalMerge: VerticalMergeType.CONTINUE,
        };

        if (!extraMap.has(ri + 1)) {
          extraMap.set(ri + 1, []);
        }
        extraMap.get(ri + 1)!.push({ cell: continueCell, columnIndex: colIdx });
      }

      colIdx += columnSpan;
    }
  }

  return extraMap;
}

// ── Descriptor ──

export const tableDesc: CustomDescriptor<TableOptions, BodyContext> = {
  kind: "custom",

  stringify(opts, ctx) {
    const parts: string[] = [];

    // Table properties
    // tblPr is required in CT_Tbl (minOccurs defaults to 1) — always emit it,
    // even when empty; do not inject optional defaults (width/borders are XSD-optional).
    const tblPrOpts: TablePropertiesOptions = {
      alignment: opts.alignment,
      borders: opts.borders,
      caption: opts.caption,
      cellMargin: opts.margins,
      cellSpacing: opts.cellSpacing,
      description: opts.description,
      float: opts.float,
      indent: opts.indent,
      layout: opts.layout,
      revision: opts.revision,
      style: opts.style,
      styleColBandSize: opts.styleColBandSize,
      styleRowBandSize: opts.styleRowBandSize,
      tableLook: opts.tableLook,
      visuallyRightToLeft: opts.visuallyRightToLeft,
      width: opts.width,
      includeIfEmpty: true,
    };
    parts.push(stringifyTableProperties(tblPrOpts)!);

    // Table grid
    const columnWidths =
      opts.columnWidths ??
      Array(Math.max(1, ...opts.rows.map((r) => ("sdt" in r ? 0 : r.cells.length)))).fill(100);
    parts.push(buildTableGridXml(columnWidths, opts.columnWidthsRevision));

    // Compute vertical merge CONTINUE cells
    const extraCells = computeVerticalMergeCells(opts.rows);

    // Rows (a row may be wrapped by a row-level SDT)
    for (let ri = 0; ri < opts.rows.length; ri++) {
      const r = opts.rows[ri];
      if ("sdt" in r) {
        const sdt = r.sdt;
        const contentXml = (sdt.rows ?? []).map((rr) => stringifyTableRow(rr, ctx)).join("");
        parts.push(stringifySdtShell(sdt.properties, sdt.endProperties, contentXml));
      } else {
        const extras = extraCells.get(ri);
        parts.push(stringifyTableRow(r, ctx, extras));
      }
    }

    return `<w:tbl>${parts.join("")}</w:tbl>`;
  },

  parse(el, ctx) {
    return parseTableEl(el, ctx as DocxReadContext);
  },
};

// ── Parse (Element → TableOptions) ──

type ParseChildFn = (el: Element, ctx: DocxReadContext) => SectionChild;

/** Callback used by table parser to parse body children. */
let _parseChild: ParseChildFn | undefined;

/** Set the child parser callback (called from parseBody). */
export function setTableParseChild(fn: ParseChildFn): void {
  _parseChild = fn;
}

function parseTablePropertiesEl(el: Element): Record<string, unknown> {
  const opts: Record<string, unknown> = {};

  const style = findChild(el, "w:tblStyle");
  if (style) {
    const val = attr(style, "w:val");
    if (val) opts.style = val;
  }

  const tblW = findChild(el, "w:tblW");
  if (tblW) {
    const rawSize = attr(tblW, "w:w");
    const type = attr(tblW, "w:type");
    const size = type === "pct" ? rawSize : attrNum(tblW, "w:w");
    if (size !== undefined || type) {
      opts.width = { size: size ?? 0, ...(type ? { type } : {}) };
    }
  }

  const jc = findChild(el, "w:jc");
  if (jc) {
    const val = attr(jc, "w:val");
    if (val) opts.alignment = val;
  }

  const layout = findChild(el, "w:tblLayout");
  if (layout) {
    const val = attr(layout, "w:type");
    if (val === "autofit" || val === "fixed") opts.layout = val;
  }

  const tblBorders = findChild(el, "w:tblBorders");
  if (tblBorders) {
    // XML side names → TableBordersOptions keys (insideH/insideV map to
    // insideHorizontal/insideVertical); all 6 sides are CT_TblBorders-optional.
    const SIDE_KEYS: ReadonlyArray<[string, keyof TableBordersOptions]> = [
      ["top", "top"],
      ["left", "left"],
      ["bottom", "bottom"],
      ["right", "right"],
      ["insideH", "insideHorizontal"],
      ["insideV", "insideVertical"],
    ];
    const borders: Partial<TableBordersOptions> = {};
    for (const [xmlSide, key] of SIDE_KEYS) {
      const sideEl = findChild(tblBorders, `w:${xmlSide}`);
      if (!sideEl) continue;
      // CT_Border requires w:val (style); skip malformed sides
      const style = attr(sideEl, "w:val");
      if (!style || !BORDER_STYLES.includes(style)) continue;
      const sideOpts: BorderOptions = { style: style as BorderOptions["style"] };
      const color = attr(sideEl, "w:color");
      if (color) sideOpts.color = color;
      const size = attrNum(sideEl, "w:sz");
      if (size !== undefined) sideOpts.size = size;
      const space = attrNum(sideEl, "w:space");
      if (space !== undefined) sideOpts.space = space;
      const themeColor = attr(sideEl, "w:themeColor");
      if (themeColor && THEME_COLORS.includes(themeColor)) {
        sideOpts.themeColor = themeColor as BorderOptions["themeColor"];
      }
      const themeTint = attr(sideEl, "w:themeTint");
      if (themeTint) sideOpts.themeTint = themeTint;
      const themeShade = attr(sideEl, "w:themeShade");
      if (themeShade) sideOpts.themeShade = themeShade;
      const shadow = attrBool(sideEl, "w:shadow");
      if (shadow !== undefined) sideOpts.shadow = shadow;
      const frame = attrBool(sideEl, "w:frame");
      if (frame !== undefined) sideOpts.frame = frame;
      borders[key] = sideOpts;
    }
    if (Object.keys(borders).length > 0) opts.borders = borders;
  }

  const tblCellMar = findChild(el, "w:tblCellMar");
  if (tblCellMar) {
    const margins: Record<string, unknown> = {};
    for (const side of ["top", "bottom", "left", "right"] as const) {
      const sideEl = findChild(tblCellMar, `w:${side}`);
      if (sideEl) {
        const size = attrNum(sideEl, "w:w");
        const type = attr(sideEl, "w:type");
        if (size !== undefined) {
          margins[side] = { size, type: type ?? "dxa" };
        }
      }
    }
    if (Object.keys(margins).length > 0) opts.margins = margins;
  }

  const shd = findChild(el, "w:shd");
  if (shd) {
    const shading: Record<string, unknown> = {};
    const fill = attr(shd, "w:fill");
    if (fill) shading.fill = fill;
    const color = attr(shd, "w:color");
    if (color) shading.color = color;
    const val = attr(shd, "w:val");
    if (val) shading.type = val;
    if (Object.keys(shading).length > 0) opts.shading = shading;
  }

  // description → w:tblDescription/@w:val
  const tblDesc = findChild(el, "w:tblDescription");
  if (tblDesc) {
    const val = attr(tblDesc, "w:val");
    if (val) opts.description = val;
  }

  // float → w:tblpPr attributes
  const tblpPr = findChild(el, "w:tblpPr");
  if (tblpPr) {
    const floatOpts: Record<string, unknown> = {};
    const horzAnchor = attr(tblpPr, "w:horzAnchor");
    if (horzAnchor) floatOpts.horizontalAnchor = horzAnchor;
    const vertAnchor = attr(tblpPr, "w:vertAnchor");
    if (vertAnchor) floatOpts.verticalAnchor = vertAnchor;
    const tblpX = attrNum(tblpPr, "w:tblpX");
    if (tblpX !== undefined) floatOpts.absoluteHorizontalPosition = tblpX;
    const tblpXSpec = attr(tblpPr, "w:tblpXSpec");
    if (tblpXSpec) floatOpts.relativeHorizontalPosition = tblpXSpec;
    const tblpY = attrNum(tblpPr, "w:tblpY");
    if (tblpY !== undefined) floatOpts.absoluteVerticalPosition = tblpY;
    const tblpYSpec = attr(tblpPr, "w:tblpYSpec");
    if (tblpYSpec) floatOpts.relativeVerticalPosition = tblpYSpec;
    const bottomFromText = attrNum(tblpPr, "w:bottomFromText");
    if (bottomFromText !== undefined) floatOpts.bottomFromText = bottomFromText;
    const topFromText = attrNum(tblpPr, "w:topFromText");
    if (topFromText !== undefined) floatOpts.topFromText = topFromText;
    const leftFromText = attrNum(tblpPr, "w:leftFromText");
    if (leftFromText !== undefined) floatOpts.leftFromText = leftFromText;
    const rightFromText = attrNum(tblpPr, "w:rightFromText");
    if (rightFromText !== undefined) floatOpts.rightFromText = rightFromText;
    if (Object.keys(floatOpts).length > 0) opts.float = floatOpts;
  }

  // indent → w:tblInd/@w:w and @w:type
  const tblInd = findChild(el, "w:tblInd");
  if (tblInd) {
    const size = attrNum(tblInd, "w:w");
    const type = attr(tblInd, "w:type");
    if (size !== undefined) {
      opts.indent = { size, ...(type ? { type } : {}) };
    }
  }

  // visuallyRightToLeft → w:bidiVisual
  const bidiVisual = findChild(el, "w:bidiVisual");
  if (bidiVisual) opts.visuallyRightToLeft = attrBool(bidiVisual, "w:val") ?? true;

  // styleRowBandSize / styleColBandSize
  const tblStyleRowBandSize = findChild(el, "w:tblStyleRowBandSize");
  if (tblStyleRowBandSize) {
    const val = attrNum(tblStyleRowBandSize, "w:val");
    if (val !== undefined) opts.styleRowBandSize = val;
  }
  const tblStyleColBandSize = findChild(el, "w:tblStyleColBandSize");
  if (tblStyleColBandSize) {
    const val = attrNum(tblStyleColBandSize, "w:val");
    if (val !== undefined) opts.styleColBandSize = val;
  }

  // caption → w:tblCaption/@w:val
  const tblCaption = findChild(el, "w:tblCaption");
  if (tblCaption) {
    const val = attr(tblCaption, "w:val");
    if (val) opts.caption = val;
  }

  // cellSpacing → w:tblCellSpacing
  const tblCellSpacing = findChild(el, "w:tblCellSpacing");
  if (tblCellSpacing) {
    const type = attr(tblCellSpacing, "w:type");
    const w = attrNum(tblCellSpacing, "w:w");
    if (w !== undefined) opts.cellSpacing = { value: w, ...(type ? { type } : {}) };
  }

  // Revision (w:tblPrChange)
  const tblPrChange = findChild(el, "w:tblPrChange");
  if (tblPrChange) {
    const rev: Record<string, unknown> = {};
    const author = attr(tblPrChange, "w:author");
    if (author) rev.author = author;
    const date = attr(tblPrChange, "w:date");
    if (date) rev.date = date;
    const id = attrNum(tblPrChange, "w:id");
    if (id !== undefined) rev.id = id;
    const innerTblPr = findChild(tblPrChange, "w:tblPr");
    if (innerTblPr) {
      Object.assign(rev, parseTablePropertiesEl(innerTblPr));
    }
    if (Object.keys(rev).length > 0) opts.revision = rev;
  }

  return opts;
}

function parseColumnWidthsEl(el: Element): number[] {
  const cols: number[] = [];
  const tblGrid = findChild(el, "w:tblGrid");
  if (!tblGrid) return cols;

  for (const col of children(tblGrid, "w:gridCol")) {
    const w = attrNum(col, "w:w");
    cols.push(w ?? 100);
  }

  return cols;
}

function parseTableRowPropertiesEl(el: Element): Record<string, unknown> {
  const opts: Record<string, unknown> = {};

  const trHeight = findChild(el, "w:trHeight");
  if (trHeight) {
    const val = attrNum(trHeight, "w:val");
    const rule = attr(trHeight, "w:hRule");
    if (val !== undefined) {
      opts.height = { value: val, rule: rule ?? "atLeast" };
    }
  }

  // cnfStyle → w:cnfStyle/@w:val (+ @w:changed)
  const cnfStyle = findChild(el, "w:cnfStyle");
  if (cnfStyle) {
    const val = attr(cnfStyle, "w:val");
    if (val) {
      const cnf: Record<string, unknown> = { val };
      const changed = attrBool(cnfStyle, "w:changed");
      if (changed !== undefined) cnf.changed = changed;
      opts.cnfStyle = cnf;
    }
  }

  // divId → w:divId/@w:val
  const divId = findChild(el, "w:divId");
  if (divId) {
    const val = attrNum(divId, "w:val");
    if (val !== undefined) opts.divId = val;
  }

  // gridBefore / gridAfter
  const gridBefore = findChild(el, "w:gridBefore");
  if (gridBefore) {
    const val = attrNum(gridBefore, "w:val");
    if (val !== undefined) opts.gridBefore = val;
  }
  const gridAfter = findChild(el, "w:gridAfter");
  if (gridAfter) {
    const val = attrNum(gridAfter, "w:val");
    if (val !== undefined) opts.gridAfter = val;
  }

  // wBefore / wAfter → widthBefore / widthAfter
  const wBefore = findChild(el, "w:wBefore");
  if (wBefore) {
    const rawSize = attr(wBefore, "w:w");
    const type = attr(wBefore, "w:type");
    const size = type === "pct" ? rawSize : attrNum(wBefore, "w:w");
    if (size !== undefined) opts.widthBefore = { size, ...(type ? { type } : {}) };
  }
  const wAfter = findChild(el, "w:wAfter");
  if (wAfter) {
    const rawSize = attr(wAfter, "w:w");
    const type = attr(wAfter, "w:type");
    const size = type === "pct" ? rawSize : attrNum(wAfter, "w:w");
    if (size !== undefined) opts.widthAfter = { size, ...(type ? { type } : {}) };
  }

  // rowAlignment → w:jc/@w:val
  const jc = findChild(el, "w:jc");
  if (jc) {
    const val = attr(jc, "w:val");
    if (val) opts.rowAlignment = val;
  }

  // hidden → w:hidden
  const hidden = findChild(el, "w:hidden");
  if (hidden) opts.hidden = attrBool(hidden, "w:val") ?? true;

  // cellSpacing → w:tblCellSpacing
  const tblCellSpacing = findChild(el, "w:tblCellSpacing");
  if (tblCellSpacing) {
    const type = attr(tblCellSpacing, "w:type");
    const w = attrNum(tblCellSpacing, "w:w");
    if (w !== undefined) opts.cellSpacing = { value: w, ...(type ? { type } : {}) };
  }

  // insertion / deletion (track changes)
  const ins = findChild(el, "w:ins");
  if (ins) opts.insertion = parseChangeAttrs(ins);
  const del = findChild(el, "w:del");
  if (del) opts.deletion = parseChangeAttrs(del);

  // Revision (w:trPrChange)
  const trPrChange = findChild(el, "w:trPrChange");
  if (trPrChange) {
    const rev: Record<string, unknown> = {};
    const author = attr(trPrChange, "w:author");
    if (author) rev.author = author;
    const date = attr(trPrChange, "w:date");
    if (date) rev.date = date;
    const id = attrNum(trPrChange, "w:id");
    if (id !== undefined) rev.id = id;
    const innerTrPr = findChild(trPrChange, "w:trPr");
    if (innerTrPr) {
      Object.assign(rev, parseTableRowPropertiesEl(innerTrPr));
    }
    if (Object.keys(rev).length > 0) opts.revision = rev;
  }

  const tblHeader = findChild(el, "w:tblHeader");
  if (tblHeader) {
    opts.tableHeader = attrBool(tblHeader, "w:val") ?? true;
  }

  const cantSplit = findChild(el, "w:cantSplit");
  if (cantSplit) {
    opts.cantSplit = attrBool(cantSplit, "w:val") ?? true;
  }

  // tblLook — conditional formatting flags (CT_TblLook)
  const tblLook = findChild(el, "w:tblLook");
  if (tblLook) {
    const look: TableLookOptions = {};
    const firstRow = attrBool(tblLook, "w:firstRow");
    if (firstRow !== undefined) look.firstRow = firstRow;
    const lastRow = attrBool(tblLook, "w:lastRow");
    if (lastRow !== undefined) look.lastRow = lastRow;
    const firstColumn = attrBool(tblLook, "w:firstColumn");
    if (firstColumn !== undefined) look.firstColumn = firstColumn;
    const lastColumn = attrBool(tblLook, "w:lastColumn");
    if (lastColumn !== undefined) look.lastColumn = lastColumn;
    const noHBand = attrBool(tblLook, "w:noHBand");
    if (noHBand !== undefined) look.noHBand = noHBand;
    const noVBand = attrBool(tblLook, "w:noVBand");
    if (noVBand !== undefined) look.noVBand = noVBand;
    if (Object.keys(look).length > 0) opts.tableLook = look;
  }

  return opts;
}

function parseTableCellPropertiesEl(el: Element): Record<string, unknown> {
  const opts: Record<string, unknown> = {};

  const tcW = findChild(el, "w:tcW");
  if (tcW) {
    const size = attrNum(tcW, "w:w");
    const type = attr(tcW, "w:type");
    if (size !== undefined) {
      opts.width = { size, type: type ?? "dxa" };
    }
  }

  const gridSpan = findChild(el, "w:gridSpan");
  if (gridSpan) {
    const val = attrNum(gridSpan, "w:val");
    if (val !== undefined) opts.columnSpan = val;
  }

  const vMerge = findChild(el, "w:vMerge");
  if (vMerge) {
    const val = attr(vMerge, "w:val");
    opts.verticalMerge = val === "restart" ? "restart" : "continue";
  }

  const vAlign = findChild(el, "w:vAlign");
  if (vAlign) {
    const val = attr(vAlign, "w:val");
    if (val) opts.verticalAlign = val;
  }

  const shd = findChild(el, "w:shd");
  if (shd) {
    const shading: Record<string, unknown> = {};
    const fill = attr(shd, "w:fill");
    if (fill) shading.fill = fill;
    const color = attr(shd, "w:color");
    if (color) shading.color = color;
    const val = attr(shd, "w:val");
    if (val) shading.type = val;
    if (Object.keys(shading).length > 0) opts.shading = shading;
  }

  const tcBorders = findChild(el, "w:tcBorders");
  if (tcBorders) {
    const borders: Record<string, unknown> = {};
    // XML side name → TableCellBordersOptions key (incl. start/end + diagonals)
    const SIDE_KEYS: ReadonlyArray<[string, string]> = [
      ["top", "top"],
      ["start", "start"],
      ["left", "left"],
      ["bottom", "bottom"],
      ["end", "end"],
      ["right", "right"],
      ["tl2br", "topLeftToBottomRight"],
      ["tr2bl", "topRightToBottomLeft"],
    ];
    for (const [xmlSide, key] of SIDE_KEYS) {
      const sideEl = findChild(tcBorders, `w:${xmlSide}`);
      if (sideEl) {
        const b: Record<string, unknown> = {};
        const val = attr(sideEl, "w:val");
        if (val) b.style = val;
        const color = attr(sideEl, "w:color");
        if (color) b.color = color;
        const sz = attrNum(sideEl, "w:sz");
        if (sz !== undefined) b.size = sz;
        borders[key] = b;
      }
    }
    if (Object.keys(borders).length > 0) opts.borders = borders;
  }

  const noWrap = findChild(el, "w:noWrap");
  if (noWrap) opts.noWrap = attrBool(noWrap, "w:val") ?? true;

  const tcMar = findChild(el, "w:tcMar");
  if (tcMar) {
    const margins: Record<string, unknown> = {};
    let marginUnitType: string | undefined;
    for (const side of ["top", "bottom", "left", "right"] as const) {
      const sideEl = findChild(tcMar, `w:${side}`);
      if (sideEl) {
        const size = attrNum(sideEl, "w:w");
        const type = attr(sideEl, "w:type");
        if (size !== undefined) {
          margins[side] = size;
          if (type && !marginUnitType) marginUnitType = type;
        }
      }
    }
    if (marginUnitType) margins.marginUnitType = marginUnitType;
    if (Object.keys(margins).length > 0) opts.margins = margins;
  }

  const textDirection = findChild(el, "w:textDirection");
  if (textDirection) {
    const val = attr(textDirection, "w:val");
    if (val) opts.textDirection = val;
  }

  // horizontalMerge → w:hMerge
  const hMerge = findChild(el, "w:hMerge");
  if (hMerge) {
    const val = attr(hMerge, "w:val");
    opts.horizontalMerge = val === "restart" ? "restart" : "continue";
  }

  // fitText → w:tcFitText
  const tcFitText = findChild(el, "w:tcFitText");
  if (tcFitText) opts.fitText = attrBool(tcFitText, "w:val") ?? true;

  // hideMark → w:hideMark
  const hideMark = findChild(el, "w:hideMark");
  if (hideMark) opts.hideMark = attrBool(hideMark, "w:val") ?? true;

  // headers → w:headers/w:header
  const headersEl = findChild(el, "w:headers");
  if (headersEl) {
    const headerVals: string[] = [];
    for (const h of headersEl.elements ?? []) {
      if (h.name !== "w:header") continue;
      const val = attr(h, "w:val");
      if (val) headerVals.push(val);
    }
    if (headerVals.length > 0) opts.headers = headerVals;
  }

  // insertion / deletion (track changes)
  const cellIns = findChild(el, "w:cellIns");
  if (cellIns) opts.insertion = parseChangeAttrs(cellIns);
  const cellDel = findChild(el, "w:cellDel");
  if (cellDel) opts.deletion = parseChangeAttrs(cellDel);

  // Revision (w:tcPrChange)
  const tcPrChange = findChild(el, "w:tcPrChange");
  if (tcPrChange) {
    const rev: Record<string, unknown> = {};
    const author = attr(tcPrChange, "w:author");
    if (author) rev.author = author;
    const date = attr(tcPrChange, "w:date");
    if (date) rev.date = date;
    const id = attrNum(tcPrChange, "w:id");
    if (id !== undefined) rev.id = id;
    const innerTcPr = findChild(tcPrChange, "w:tcPr");
    if (innerTcPr) {
      Object.assign(rev, parseTableCellPropertiesEl(innerTcPr));
    }
    if (Object.keys(rev).length > 0) opts.revision = rev;
  }

  // cellMerge → w:cellMerge
  const cellMerge = findChild(el, "w:cellMerge");
  if (cellMerge) {
    const cm = parseChangeAttrs(cellMerge);
    const vMerge = attr(cellMerge, "w:vMerge");
    if (vMerge) cm.verticalMerge = xsdVerticalMergeRev.from(vMerge) as "continue" | "restart";
    const vMergeOrig = attr(cellMerge, "w:vMergeOrig");
    if (vMergeOrig) {
      cm.verticalMergeOriginal = xsdVerticalMergeRev.from(vMergeOrig) as "continue" | "restart";
    }
    if (Object.keys(cm).length > 0) opts.cellMerge = cm;
  }

  return opts;
}

function parseTableCellEl(el: Element, ctx: DocxReadContext): TableCellOptions {
  const opts: Record<string, unknown> = {};

  const tcPr = findChild(el, "w:tcPr");
  if (tcPr) {
    Object.assign(opts, parseTableCellPropertiesEl(tcPr));
  }

  const childElements: SectionChild[] = [];
  for (const child of el.elements ?? []) {
    switch (child.name) {
      case "w:tcPr":
        break;
      case "w:p":
      case "w:tbl":
        if (_parseChild) childElements.push(_parseChild(child, ctx));
        break;
      default:
        break;
    }
  }

  opts.children = childElements;
  return opts as unknown as TableCellOptions;
}

function parseTableRowEl(el: Element, ctx: DocxReadContext): TableRowOptions {
  const opts: Record<string, unknown> = {};

  const trPr = findChild(el, "w:trPr");
  if (trPr) {
    Object.assign(opts, parseTableRowPropertiesEl(trPr));
  }

  // rsid attributes on w:tr element
  for (const [attrName, optKey] of [
    ["w:rsidRPr", "rsidRPr"],
    ["w:rsidR", "rsidR"],
    ["w:rsidDel", "rsidDel"],
    ["w:rsidTr", "rsidTr"],
  ] as const) {
    const val = attr(el, attrName);
    if (val) opts[optKey] = val;
  }

  const childCells: (TableCellOptions | { sdt: SdtCellOptions })[] = [];
  for (const child of el.elements ?? []) {
    if (child.name === "w:tc") {
      childCells.push(parseTableCellEl(child, ctx));
    } else if (child.name === "w:sdt") {
      const sdtPr = findChild(child, "w:sdtPr");
      const properties = sdtPr ? parseSdtProperties(sdtPr) : {};
      const sdtEndPr = findChild(child, "w:sdtEndPr");
      const endProperties = sdtEndPr ? parseRunProperties(sdtEndPr) : undefined;
      const sdtContent = findChild(child, "w:sdtContent");
      const sdtCells: TableCellOptions[] = [];
      if (sdtContent) {
        for (const sub of sdtContent.elements ?? []) {
          if (sub.name === "w:tc") sdtCells.push(parseTableCellEl(sub, ctx));
        }
      }
      const sdt: SdtCellOptions = {
        properties,
      };
      if (sdtCells.length > 0) sdt.cells = sdtCells;
      if (endProperties) sdt.endProperties = endProperties;
      childCells.push({ sdt });
    }
  }

  opts.cells = childCells;
  return opts as unknown as TableRowOptions;
}

function parseTableEl(el: Element, ctx: DocxReadContext): TableOptions {
  const opts: Record<string, unknown> = {};

  const tblPr = findChild(el, "w:tblPr");
  if (tblPr) {
    Object.assign(opts, parseTablePropertiesEl(tblPr));
  }

  const colWidths = parseColumnWidthsEl(el);
  if (colWidths.length > 0) {
    opts.columnWidths = colWidths;
  }

  const rows: (TableRowOptions | { sdt: SdtRowOptions })[] = [];
  for (const child of el.elements ?? []) {
    if (child.name === "w:tr") {
      rows.push(parseTableRowEl(child, ctx));
    } else if (child.name === "w:sdt") {
      const sdtPr = findChild(child, "w:sdtPr");
      const properties = sdtPr ? parseSdtProperties(sdtPr) : {};
      const sdtEndPr = findChild(child, "w:sdtEndPr");
      const endProperties = sdtEndPr ? parseRunProperties(sdtEndPr) : undefined;
      const sdtContent = findChild(child, "w:sdtContent");
      const sdtRows: TableRowOptions[] = [];
      if (sdtContent) {
        for (const sub of sdtContent.elements ?? []) {
          if (sub.name === "w:tr") sdtRows.push(parseTableRowEl(sub, ctx));
        }
      }
      const sdt: SdtRowOptions = {
        properties,
      };
      if (sdtRows.length > 0) sdt.rows = sdtRows;
      if (endProperties) sdt.endProperties = endProperties;
      rows.push({ sdt });
    }
  }

  opts.rows = rows;
  return opts as unknown as TableOptions;
}
