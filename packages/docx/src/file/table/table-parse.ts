import type { SectionChild } from "@file/section-child";
import type { ITableOptions } from "@file/table/table";
import type { ITableCellOptions } from "@file/table/table-cell";
import type { ITableRowOptions } from "@file/table/table-row";
/**
 * Table parser for DOCX documents.
 *
 * Parses w:tbl, w:tr, w:tc Element trees into ITableOptions.
 *
 * @module
 */
import { attr, attrBool, attrNum, children, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { ParseContext } from "../../parse/context";
import { parseParagraph } from "../paragraph/paragraph-parse";

/**
 * Parse w:tblPr into table properties.
 */
function parseTableProperties(el: Element): Record<string, unknown> {
  const opts: Record<string, unknown> = {};

  const style = findChild(el, "w:tblStyle");
  if (style) {
    const val = attr(style, "w:val");
    if (val) opts.style = val;
  }

  const tblW = findChild(el, "w:tblW");
  if (tblW) {
    const size = attrNum(tblW, "w:w");
    const type = attr(tblW, "w:type");
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

  // Borders
  const tblBorders = findChild(el, "w:tblBorders");
  if (tblBorders) {
    const borders: Record<string, unknown> = {};
    for (const side of ["top", "bottom", "left", "right", "insideH", "insideV"] as const) {
      const sideEl = findChild(tblBorders, `w:${side}`);
      if (sideEl) {
        const b: Record<string, unknown> = {};
        const val = attr(sideEl, "w:val");
        if (val) b.style = val;
        const color = attr(sideEl, "w:color");
        if (color) b.color = color;
        const sz = attrNum(sideEl, "w:sz");
        if (sz !== undefined) b.size = sz;
        const space = attrNum(sideEl, "w:space");
        if (space !== undefined) b.space = space;
        borders[side] = b;
      }
    }
    if (Object.keys(borders).length > 0) opts.borders = borders;
  }

  // Cell margin
  const tblCellMar = findChild(el, "w:tblCellMar");
  if (tblCellMar) {
    const margins: Record<string, unknown> = {};
    for (const side of ["top", "bottom", "left", "right"] as const) {
      const sideEl = findChild(tblCellMar, `w:${side}`);
      if (sideEl) {
        const size = attrNum(sideEl, "w:w");
        const type = attr(sideEl, "w:type");
        if (size !== undefined) {
          (margins as Record<string, unknown>)[side] = { size, type: type ?? "dxa" };
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

  return opts;
}

/**
 * Parse w:gridCol elements from w:tblGrid to extract column widths.
 */
function parseColumnWidths(el: Element): number[] {
  const cols: number[] = [];
  const tblGrid = findChild(el, "w:tblGrid");
  if (!tblGrid) return cols;

  for (const col of children(tblGrid, "w:gridCol")) {
    const w = attrNum(col, "w:w");
    cols.push(w ?? 100);
  }

  return cols;
}

/**
 * Parse w:trPr into table row properties.
 */
function parseTableRowProperties(el: Element): Record<string, unknown> {
  const opts: Record<string, unknown> = {};

  const trHeight = findChild(el, "w:trHeight");
  if (trHeight) {
    const val = attrNum(trHeight, "w:val");
    const rule = attr(trHeight, "w:hRule");
    if (val !== undefined) {
      opts.height = { value: val, rule: rule ?? "atLeast" };
    }
  }

  const tblHeader = findChild(el, "w:tblHeader");
  if (tblHeader) {
    opts.tableHeader = attrBool(tblHeader, "w:val") ?? true;
  }

  const cantSplit = findChild(el, "w:cantSplit");
  if (cantSplit) {
    opts.cantSplit = attrBool(cantSplit, "w:val") ?? true;
  }

  return opts;
}

/**
 * Parse w:tcPr into table cell properties.
 */
function parseTableCellProperties(el: Element): Record<string, unknown> {
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

  // Cell borders
  const tcBorders = findChild(el, "w:tcBorders");
  if (tcBorders) {
    const borders: Record<string, unknown> = {};
    for (const side of ["top", "bottom", "left", "right"] as const) {
      const sideEl = findChild(tcBorders, `w:${side}`);
      if (sideEl) {
        const b: Record<string, unknown> = {};
        const val = attr(sideEl, "w:val");
        if (val) b.style = val;
        const color = attr(sideEl, "w:color");
        if (color) b.color = color;
        const sz = attrNum(sideEl, "w:sz");
        if (sz !== undefined) b.size = sz;
        borders[side] = b;
      }
    }
    if (Object.keys(borders).length > 0) opts.borders = borders;
  }

  const noWrap = findChild(el, "w:noWrap");
  if (noWrap) opts.noWrap = true;

  const textDirection = findChild(el, "w:textDirection");
  if (textDirection) {
    const val = attr(textDirection, "w:val");
    if (val) opts.textDirection = val;
  }

  return opts;
}

// Forward declaration to handle circular reference
let _parseSectionChild: ((el: Element, ctx: ParseContext) => SectionChild) | undefined;

/**
 * Set the section child parser function (called by parse-body to resolve circular dependency).
 */
export function setSectionChildParser(fn: (el: Element, ctx: ParseContext) => SectionChild): void {
  _parseSectionChild = fn;
}

/**
 * Parse w:tc element into ITableCellOptions.
 */
export function parseTableCell(el: Element, ctx: ParseContext): ITableCellOptions {
  const opts: Record<string, unknown> = {};

  const tcPr = findChild(el, "w:tcPr");
  if (tcPr) {
    Object.assign(opts, parseTableCellProperties(tcPr));
  }

  const childElements: SectionChild[] = [];
  for (const child of el.elements ?? []) {
    switch (child.name) {
      case "w:tcPr":
        break;
      case "w:p":
        if (_parseSectionChild) {
          childElements.push(_parseSectionChild(child, ctx));
        } else {
          childElements.push({ paragraph: parseParagraph(child, ctx) });
        }
        break;
      case "w:tbl":
        if (_parseSectionChild) {
          childElements.push(_parseSectionChild(child, ctx));
        } else {
          childElements.push({ table: parseTable(child, ctx) });
        }
        break;
      default:
        break;
    }
  }

  opts.children = childElements;
  return opts as unknown as ITableCellOptions;
}

/**
 * Parse w:tr element into ITableRowOptions.
 */
export function parseTableRow(el: Element, ctx: ParseContext): ITableRowOptions {
  const opts: Record<string, unknown> = {};

  const trPr = findChild(el, "w:trPr");
  if (trPr) {
    Object.assign(opts, parseTableRowProperties(trPr));
  }

  const childCells: ITableCellOptions[] = [];
  for (const child of el.elements ?? []) {
    if (child.name === "w:tc") {
      childCells.push(parseTableCell(child, ctx));
    }
  }

  opts.children = childCells;
  return opts as unknown as ITableRowOptions;
}

/**
 * Parse w:tbl element into ITableOptions.
 */
export function parseTable(el: Element, ctx: ParseContext): ITableOptions {
  const opts: Record<string, unknown> = {};

  const tblPr = findChild(el, "w:tblPr");
  if (tblPr) {
    Object.assign(opts, parseTableProperties(tblPr));
  }

  // Column widths
  const colWidths = parseColumnWidths(el);
  if (colWidths.length > 0) {
    opts.columnWidths = colWidths;
  }

  // Rows
  const rows: ITableRowOptions[] = [];
  for (const child of el.elements ?? []) {
    if (child.name === "w:tr") {
      rows.push(parseTableRow(child, ctx));
    }
  }

  opts.rows = rows;
  return opts as unknown as ITableOptions;
}
