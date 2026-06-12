/**
 * Table (w:tbl) descriptor for DOCX.
 *
 * Stringifies pure JSON TableOptions into XML using direct string
 * concatenation — zero IXmlableObject, zero xml(), zero BaseXmlComponent.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrBool, attrNum, children, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { stringifyParagraphInline } from "@parts/inline";
import type { TableGridChangeOptions } from "@parts/table/grid";
import type { TableOptions } from "@parts/table/table";
import type { TableCellOptions } from "@parts/table/table-cell/table-cell";
import { VerticalMergeType } from "@parts/table/table-cell/table-cell-components";
import type { TableRowOptions } from "@parts/table/table-row/table-row";
import type { SectionChild } from "@shared/section";

import type { BodyContext } from "../../context";
import {
  stringifyTableCellProperties,
  stringifyTableProperties,
  stringifyTablePropertyExceptions,
  stringifyTableRowProperties,
  type TablePropertiesOptions,
} from "./stringify";

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

  // Cells
  for (const cell of row.cells) {
    parts.push(stringifyTableCell(cell as TableCellOptions, ctx));
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
    const { columnSpan } = getCellSpans(cells[i] as TableCellOptions);
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
  rows: TableRowOptions[],
): Map<number, { cell: TableCellOptions; columnIndex: number }[]> {
  const extraMap = new Map<number, { cell: TableCellOptions; columnIndex: number }[]>();
  for (let ri = 0; ri < rows.length - 1; ri++) {
    const cells = rows[ri].cells;
    let colIdx = 0;

    for (const cell of cells) {
      const typedCell = cell as TableCellOptions;
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
    const tblPrOpts: TablePropertiesOptions = {
      alignment: opts.alignment,
      borders: opts.borders ?? {},
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
      width: opts.width ?? { size: 100 },
    };
    const tblPr = stringifyTableProperties(tblPrOpts);
    if (tblPr) parts.push(tblPr);

    // Table grid
    const columnWidths =
      opts.columnWidths ?? Array(Math.max(...opts.rows.map((r) => r.cells.length))).fill(100);
    parts.push(buildTableGridXml(columnWidths, opts.columnWidthsRevision));

    // Compute vertical merge CONTINUE cells
    const extraCells = computeVerticalMergeCells(opts.rows as TableRowOptions[]);

    // Rows
    for (let ri = 0; ri < opts.rows.length; ri++) {
      const row = opts.rows[ri] as TableRowOptions;
      const extras = extraCells.get(ri);
      parts.push(stringifyTableRow(row, ctx, extras));
    }

    return `<w:tbl>${parts.join("")}</w:tbl>`;
  },

  parse(el, ctx) {
    return parseTableEl(el, ctx as import("../../context").DocxReadContext);
  },
};

// ── Parse (Element → TableOptions) ──

type ParseChildFn = (
  el: Element,
  ctx: import("../../context").DocxReadContext,
) => import("@shared/section").SectionChild;

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

  return opts;
}

function parseTableCellEl(
  el: Element,
  ctx: import("../../context").DocxReadContext,
): TableCellOptions {
  const opts: Record<string, unknown> = {};

  const tcPr = findChild(el, "w:tcPr");
  if (tcPr) {
    Object.assign(opts, parseTableCellPropertiesEl(tcPr));
  }

  const childElements: import("@shared/section").SectionChild[] = [];
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

function parseTableRowEl(
  el: Element,
  ctx: import("../../context").DocxReadContext,
): TableRowOptions {
  const opts: Record<string, unknown> = {};

  const trPr = findChild(el, "w:trPr");
  if (trPr) {
    Object.assign(opts, parseTableRowPropertiesEl(trPr));
  }

  const childCells: TableCellOptions[] = [];
  for (const child of el.elements ?? []) {
    if (child.name === "w:tc") {
      childCells.push(parseTableCellEl(child, ctx));
    }
  }

  opts.cells = childCells;
  return opts as unknown as TableRowOptions;
}

function parseTableEl(el: Element, ctx: import("../../context").DocxReadContext): TableOptions {
  const opts: Record<string, unknown> = {};

  const tblPr = findChild(el, "w:tblPr");
  if (tblPr) {
    Object.assign(opts, parseTablePropertiesEl(tblPr));
  }

  const colWidths = parseColumnWidthsEl(el);
  if (colWidths.length > 0) {
    opts.columnWidths = colWidths;
  }

  const rows: TableRowOptions[] = [];
  for (const child of el.elements ?? []) {
    if (child.name === "w:tr") {
      rows.push(parseTableRowEl(child, ctx));
    }
  }

  opts.rows = rows;
  return opts as unknown as TableOptions;
}
