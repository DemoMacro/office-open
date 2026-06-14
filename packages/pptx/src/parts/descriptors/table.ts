/**
 * Table (p:graphicFrame with a:tbl) descriptor for PPTX.
 *
 * @module
 */

import { convertPixelsToEmu, xsdTextAnchor } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { parse, stringify } from "@office-open/core/descriptor";
import type { ReadContext } from "@office-open/core/descriptor";
import { fillDesc } from "@office-open/core/drawingml";
import type { FillOptions } from "@office-open/core/drawingml";
import { attr, attrBool, attrNum, children, findChild, textOf } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { escapeXml } from "@office-open/xml";

import type { PptxWriteContext } from "../../context";
import { readPositionFromXfrm } from "./shape";
import { paragraphDesc, type ParagraphDescriptorOptions } from "./text";

// ── Types ──

export interface CellBorderDescriptorOptions {
  width?: number;
  color?: string;
  dashStyle?: "solid" | "dash" | "dashDot" | "lgDash" | "sysDot" | "sysDash";
}

export interface TableCellDescriptorOptions {
  text?: string;
  children?: (ParagraphDescriptorOptions | string)[];
  fill?: FillOptions;
  borders?: {
    top?: CellBorderDescriptorOptions;
    bottom?: CellBorderDescriptorOptions;
    left?: CellBorderDescriptorOptions;
    right?: CellBorderDescriptorOptions;
  };
  columnSpan?: number;
  rowSpan?: number;
  horizontalMerge?: "continue" | "restart";
  verticalMerge?: "continue" | "restart";
  verticalAlign?: "top" | "center" | "bottom" | "justify" | "distribute";
  margins?: {
    top?: number;
    bottom?: number;
    left?: number;
    right?: number;
  };
}

export interface TableRowDescriptorOptions {
  height?: number;
  cells: TableCellDescriptorOptions[];
}

export interface TableDescriptorOptions {
  id?: number;
  name?: string;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  rows: TableRowDescriptorOptions[];
  columnWidths?: number[];
  firstRow?: boolean;
  lastRow?: boolean;
  bandRow?: boolean;
  firstCol?: boolean;
  lastCol?: boolean;
  bandCol?: boolean;
  tableStyleId?: string;
  borders?: {
    top?: CellBorderDescriptorOptions;
    bottom?: CellBorderDescriptorOptions;
    left?: CellBorderDescriptorOptions;
    right?: CellBorderDescriptorOptions;
  };
}

// ── ID counter ──

let _nextTableId = 1024;

// ── Table (p:graphicFrame) descriptor ──

export const tableDesc: CustomDescriptor<TableDescriptorOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const pptxCtx = ctx as unknown as PptxWriteContext;
    const id = opts.id ?? _nextTableId++;
    const name = opts.name ?? `Table ${id}`;

    const x = convertPixelsToEmu(opts.x ?? 0);
    const y = convertPixelsToEmu(opts.y ?? 0);
    const w = convertPixelsToEmu(opts.width ?? 100);
    const h = convertPixelsToEmu(opts.height ?? 100);

    const parts: string[] = [];

    // p:nvGraphicFramePr
    parts.push(
      `<p:nvGraphicFramePr><p:cNvPr id="${id}" name="${escapeXml(name)}"/>` +
        `<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>` +
        `<p:nvPr/></p:nvGraphicFramePr>`,
    );

    // p:xfrm
    parts.push(`<p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></p:xfrm>`);

    // a:graphic > a:graphicData > a:tbl
    const tblParts: string[] = [];

    // a:tblPr
    tblParts.push(stringifyTblPr(opts));

    // a:tblGrid
    const colWidths =
      opts.columnWidths && opts.columnWidths.length > 0
        ? opts.columnWidths
        : Array.from({ length: opts.rows[0]?.cells.length ?? 1 }, () => 0);
    const gridCols = colWidths.map((cw) => `<a:gridCol w="${cw}"/>`).join("");
    tblParts.push(`<a:tblGrid>${gridCols}</a:tblGrid>`);

    // a:tr[] — with border distribution
    const rowCount = opts.rows.length;
    for (let ri = 0; ri < rowCount; ri++) {
      const row = opts.rows[ri];
      const cells = distributeBorders(row, ri, rowCount, opts.borders);
      tblParts.push(stringifyRow({ ...row, cells }, pptxCtx));
    }

    const tblXml = `<a:tbl>${tblParts.join("")}</a:tbl>`;
    parts.push(
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">${tblXml}</a:graphicData></a:graphic>`,
    );

    return `<p:graphicFrame>${parts.join("")}</p:graphicFrame>`;
  },

  parse(el, _ctx) {
    const result: Record<string, unknown> = {};

    // Position from p:xfrm
    const xfrm = findChild(el, "p:xfrm");
    if (xfrm) Object.assign(result, readPositionFromXfrm(xfrm));

    // Name from p:nvGraphicFramePr → p:cNvPr
    const nvGfxFramePr = findChild(el, "p:nvGraphicFramePr");
    if (nvGfxFramePr) {
      const cNvPr = findChild(nvGfxFramePr, "p:cNvPr");
      if (cNvPr) {
        const id = attrNum(cNvPr, "id");
        if (id !== undefined) result.id = id;
        const name = attr(cNvPr, "name");
        if (name) result.name = name;
      }
    }

    // Find a:tbl inside a:graphicData
    const graphicData = findChild(el, "a:graphic") ? findChild(el, "a:graphic") : undefined;
    const gd = graphicData ? findChild(graphicData, "a:graphicData") : undefined;
    const tbl = gd ? findChild(gd, "a:tbl") : undefined;
    if (!tbl) return result as unknown as TableDescriptorOptions;

    // a:tblPr
    const tblPr = findChild(tbl, "a:tblPr");
    if (tblPr) {
      if (attrBool(tblPr, "firstRow")) result.firstRow = true;
      if (attrBool(tblPr, "lastRow")) result.lastRow = true;
      if (attrBool(tblPr, "bandRow")) result.bandRow = true;
      if (attrBool(tblPr, "firstCol")) result.firstCol = true;
      if (attrBool(tblPr, "lastCol")) result.lastCol = true;
      if (attrBool(tblPr, "bandCol")) result.bandCol = true;

      const tableStyleIdEl = findChild(tblPr, "a:tableStyleId");
      if (tableStyleIdEl) {
        const styleId = textOf(tableStyleIdEl);
        if (styleId) result.tableStyleId = styleId;
      }

      // Table-level borders
      const borders: Record<string, unknown> = {};
      for (const [elName, key] of [
        ["a:lnL", "left"],
        ["a:lnR", "right"],
        ["a:lnT", "top"],
        ["a:lnB", "bottom"],
      ] as const) {
        const borderEl = findChild(tblPr, elName);
        if (borderEl) {
          const borderOpts: Record<string, unknown> = {};
          const w = attrNum(borderEl, "w");
          if (w !== undefined) borderOpts.width = w;
          const fillResult = parseTableCellFill(borderEl);
          if (fillResult) borderOpts.color = fillResult;
          const prstDash = findChild(borderEl, "a:prstDash");
          if (prstDash) {
            const val = attr(prstDash, "val");
            if (val) borderOpts.dashStyle = val;
          }
          if (Object.keys(borderOpts).length > 0) borders[key] = borderOpts;
        }
      }
      if (Object.keys(borders).length > 0) result.borders = borders;
    }

    // a:tblGrid → columnWidths
    const tblGrid = findChild(tbl, "a:tblGrid");
    if (tblGrid) {
      const colWidths: number[] = [];
      for (const gridCol of children(tblGrid, "a:gridCol")) {
        const w = attrNum(gridCol, "w");
        colWidths.push(w ?? 0);
      }
      if (colWidths.length > 0) result.columnWidths = colWidths;
    }

    // a:tr → rows
    const rows: Record<string, unknown>[] = [];
    for (const tr of children(tbl, "a:tr")) {
      const rowOpts: Record<string, unknown> = {};
      const h = attrNum(tr, "h");
      if (h !== undefined) rowOpts.height = h;

      const cells: Record<string, unknown>[] = [];
      for (const tc of children(tr, "a:tc")) {
        cells.push(parseTableCell(tc, _ctx));
      }
      rowOpts.cells = cells;
      rows.push(rowOpts);
    }
    result.rows = rows;

    return result as unknown as TableDescriptorOptions;
  },
};

// ── Helpers ──

function stringifyTblPr(opts: TableDescriptorOptions): string {
  const attrs: string[] = [];
  if (opts.firstRow !== undefined) attrs.push(`firstRow="${opts.firstRow ? 1 : 0}"`);
  if (opts.lastRow !== undefined) attrs.push(`lastRow="${opts.lastRow ? 1 : 0}"`);
  if (opts.bandRow !== undefined) attrs.push(`bandRow="${opts.bandRow ? 1 : 0}"`);
  if (opts.firstCol !== undefined) attrs.push(`firstCol="${opts.firstCol ? 1 : 0}"`);
  if (opts.lastCol !== undefined) attrs.push(`lastCol="${opts.lastCol ? 1 : 0}"`);
  if (opts.bandCol !== undefined) attrs.push(`bandCol="${opts.bandCol ? 1 : 0}"`);
  if (attrs.length === 0 && !opts.tableStyleId) return "<a:tblPr/>";
  const styleId = opts.tableStyleId
    ? `<a:tableStyleId>${escapeXml(opts.tableStyleId)}</a:tableStyleId>`
    : "";
  return `<a:tblPr ${attrs.join(" ")}>${styleId}</a:tblPr>`;
}

function stringifyRow(row: TableRowDescriptorOptions, ctx: PptxWriteContext): string {
  const h = row.height ?? 0;
  const cellParts: string[] = [];
  for (const cell of row.cells) {
    cellParts.push(stringifyCell(cell, ctx));
  }
  return `<a:tr h="${h}">${cellParts.join("")}</a:tr>`;
}

function stringifyCell(cell: TableCellDescriptorOptions, ctx: PptxWriteContext): string {
  const parts: string[] = [];

  // Attributes
  const tcAttrs: string[] = [];
  if (cell.columnSpan !== undefined && cell.columnSpan > 1)
    tcAttrs.push(`gridSpan="${cell.columnSpan}"`);
  if (cell.rowSpan !== undefined && cell.rowSpan > 1) tcAttrs.push(`rowSpan="${cell.rowSpan}"`);
  if (cell.horizontalMerge) tcAttrs.push(`hMerge="${cell.horizontalMerge === "restart" ? 1 : 0}"`);
  if (cell.verticalMerge) tcAttrs.push(`vMerge="${cell.verticalMerge === "restart" ? 1 : 0}"`);
  const tcAttrStr = tcAttrs.length > 0 ? ` ${tcAttrs.join(" ")}` : "";

  // a:txBody
  parts.push(stringifyTxBody(cell, ctx));

  // a:tcPr
  parts.push(stringifyTcPr(cell, ctx));

  return `<a:tc${tcAttrStr}>${parts.join("")}</a:tc>`;
}

function stringifyTxBody(cell: TableCellDescriptorOptions, ctx: PptxWriteContext): string {
  const txParts: string[] = [];

  // a:bodyPr
  const bodyPrAttrs: string[] = [];
  const margins = cell.margins;
  if (margins?.top !== undefined) bodyPrAttrs.push(`tIns="${margins.top}"`);
  if (margins?.bottom !== undefined) bodyPrAttrs.push(`bIns="${margins.bottom}"`);
  if (margins?.left !== undefined) bodyPrAttrs.push(`lIns="${margins.left}"`);
  if (margins?.right !== undefined) bodyPrAttrs.push(`rIns="${margins.right}"`);
  const bodyPrStr = bodyPrAttrs.length > 0 ? ` ${bodyPrAttrs.join(" ")}` : "";
  txParts.push(`<a:bodyPr${bodyPrStr}/>`);
  txParts.push("<a:lstStyle/>");

  // Paragraphs
  if (cell.children) {
    for (const c of cell.children) {
      if (typeof c === "string") {
        const pXml = paragraphDesc.stringify({ children: [{ text: c }] }, ctx);
        if (pXml) txParts.push(pXml);
      } else {
        const pXml = paragraphDesc.stringify(c, ctx);
        if (pXml) txParts.push(pXml);
      }
    }
  } else if (cell.text !== undefined) {
    const pXml = paragraphDesc.stringify({ children: [{ text: cell.text }] }, ctx);
    if (pXml) txParts.push(pXml);
  } else {
    txParts.push("<a:p/>");
  }

  return `<a:txBody>${txParts.join("")}</a:txBody>`;
}

function stringifyTcPr(cell: TableCellDescriptorOptions, ctx: PptxWriteContext): string {
  const parts: string[] = [];

  if (cell.verticalAlign) {
    parts.push(`anchor="${xsdTextAnchor.to(cell.verticalAlign)}"`);
  }

  if (cell.borders) {
    if (cell.borders.left) parts.push(buildBorderLine("a:lnL", cell.borders.left));
    if (cell.borders.right) parts.push(buildBorderLine("a:lnR", cell.borders.right));
    if (cell.borders.top) parts.push(buildBorderLine("a:lnT", cell.borders.top));
    if (cell.borders.bottom) parts.push(buildBorderLine("a:lnB", cell.borders.bottom));
  }

  if (cell.fill !== undefined) {
    const fillXml = stringify(fillDesc, cell.fill, ctx);
    if (fillXml) parts.push(fillXml);
  }

  if (parts.length === 0) return "<a:tcPr/>";

  const firstIsAttr = !parts[0].startsWith("<");
  if (firstIsAttr) {
    const attrStr = parts[0];
    const children = parts.slice(1);
    if (children.length === 0) return `<a:tcPr ${attrStr}/>`;
    return `<a:tcPr ${attrStr}>${children.join("")}</a:tcPr>`;
  }

  return `<a:tcPr>${parts.join("")}</a:tcPr>`;
}

function buildBorderLine(name: string, options: CellBorderDescriptorOptions): string {
  const attrs: string[] = [];
  if (options.width !== undefined) attrs.push(`w="${options.width}"`);

  const children: string[] = [];
  if (options.color) {
    children.push(
      `<a:solidFill><a:srgbClr val="${options.color.replace("#", "")}"/></a:solidFill>`,
    );
  }
  if (options.dashStyle) {
    children.push(`<a:prstDash val="${options.dashStyle}"/>`);
  }

  const attrStr = attrs.length > 0 ? ` ${attrs.join(" ")}` : "";
  if (children.length === 0) return `<${name}${attrStr}/>`;
  return `<${name}${attrStr}>${children.join("")}</${name}>`;
}

/** Distribute table-level borders to edge cells. */
function distributeBorders(
  row: TableRowDescriptorOptions,
  ri: number,
  rowCount: number,
  tb: TableDescriptorOptions["borders"],
): TableCellDescriptorOptions[] {
  if (!tb) return row.cells;
  const colCount = row.cells.length;
  return row.cells.map((cell, ci) => {
    const needTop = ri === 0 && !!tb.top && !cell.borders?.top;
    const needBottom = ri === rowCount - 1 && !!tb.bottom && !cell.borders?.bottom;
    const needLeft = ci === 0 && !!tb.left && !cell.borders?.left;
    const needRight = ci === colCount - 1 && !!tb.right && !cell.borders?.right;
    if (!needTop && !needBottom && !needLeft && !needRight) return cell;
    const borders = {
      ...cell.borders,
      ...(needTop && { top: tb.top }),
      ...(needBottom && { bottom: tb.bottom }),
      ...(needLeft && { left: tb.left }),
      ...(needRight && { right: tb.right }),
    };
    return { ...cell, borders };
  });
}

/** Parse a table cell (a:tc) into options. */
function parseTableCell(tc: Element, readCtx?: ReadContext): Record<string, unknown> {
  const result: Record<string, unknown> = {};
  const ctx = readCtx ?? ({} as ReadContext);

  const gridSpan = attrNum(tc, "gridSpan");
  if (gridSpan !== undefined && gridSpan > 1) result.columnSpan = gridSpan;
  const rowSpan = attrNum(tc, "rowSpan");
  if (rowSpan !== undefined && rowSpan > 1) result.rowSpan = rowSpan;
  const hMerge = attr(tc, "hMerge");
  if (hMerge === "1") result.horizontalMerge = "restart";
  else if (hMerge === "0") result.horizontalMerge = "continue";
  const vMerge = attr(tc, "vMerge");
  if (vMerge === "1") result.verticalMerge = "restart";
  else if (vMerge === "0") result.verticalMerge = "continue";

  // a:txBody → paragraph children
  const txBody = findChild(tc, "a:txBody");
  if (txBody) {
    const paragraphs: Record<string, unknown>[] = [];
    for (const pEl of txBody.elements ?? []) {
      if (pEl.name !== "a:p") continue;
      const para = paragraphDesc.parse(pEl, ctx) as Record<string, unknown>;
      paragraphs.push(para);
    }

    if (paragraphs.length === 1) {
      const p = paragraphs[0];
      if (p.text && !p.children) {
        result.text = p.text;
      } else {
        result.children = [p];
      }
    } else if (paragraphs.length > 0) {
      result.children = paragraphs;
    }

    // Extract margins from a:bodyPr
    const bodyPr = findChild(txBody, "a:bodyPr");
    if (bodyPr) {
      const margins: Record<string, number> = {};
      const tIns = attrNum(bodyPr, "tIns");
      if (tIns !== undefined) margins.top = tIns;
      const bIns = attrNum(bodyPr, "bIns");
      if (bIns !== undefined) margins.bottom = bIns;
      const lIns = attrNum(bodyPr, "lIns");
      if (lIns !== undefined) margins.left = lIns;
      const rIns = attrNum(bodyPr, "rIns");
      if (rIns !== undefined) margins.right = rIns;
      if (Object.keys(margins).length > 0) result.margins = margins;
    }
  }

  // a:tcPr
  const tcPr = findChild(tc, "a:tcPr");
  if (tcPr) {
    const anchor = attr(tcPr, "anchor");
    if (anchor) result.verticalAlign = xsdTextAnchor.from(anchor);

    // Margins from tcPr attributes
    const margins: Record<string, number> = {};
    const marL = attrNum(tcPr, "marL");
    if (marL !== undefined) margins.left = marL;
    const marR = attrNum(tcPr, "marR");
    if (marR !== undefined) margins.right = marR;
    const marT = attrNum(tcPr, "marT");
    if (marT !== undefined) margins.top = marT;
    const marB = attrNum(tcPr, "marB");
    if (marB !== undefined) margins.bottom = marB;
    if (Object.keys(margins).length > 0) result.margins = margins;

    // Fill — use full fillDesc for all fill types
    const fillResult = parse(fillDesc, tcPr, ctx);
    if (fillResult && Object.keys(fillResult).length > 0) result.fill = fillResult;

    // Cell borders
    const borders: Record<string, unknown> = {};
    for (const [elName, key] of [
      ["a:lnL", "left"],
      ["a:lnR", "right"],
      ["a:lnT", "top"],
      ["a:lnB", "bottom"],
    ] as const) {
      const borderEl = findChild(tcPr, elName);
      if (borderEl) {
        const borderOpts: Record<string, unknown> = {};
        const w = attrNum(borderEl, "w");
        if (w !== undefined) borderOpts.width = w;
        const fillResult = parseTableCellFill(borderEl);
        if (fillResult) borderOpts.color = fillResult;
        // Dash style
        const prstDash = findChild(borderEl, "a:prstDash");
        if (prstDash) {
          const val = attr(prstDash, "val");
          if (val) borderOpts.dashStyle = val;
        }
        if (Object.keys(borderOpts).length > 0) borders[key] = borderOpts;
      }
    }
    if (Object.keys(borders).length > 0) result.borders = borders;
  }

  return result;
}

/** Extract a solid fill color string from an element (e.g. a:lnL, a:tcPr). */
function parseTableCellFill(el: Element): string | undefined {
  const solidFill = findChild(el, "a:solidFill");
  if (!solidFill) return undefined;
  const srgbClr = findChild(solidFill, "a:srgbClr");
  return srgbClr?.attributes?.["val"] ? String(srgbClr.attributes["val"]) : undefined;
}
