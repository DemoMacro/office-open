/**
 * Table (p:graphicFrame with a:tbl) descriptor for PPTX.
 *
 * @module
 */

import {
  convertToEmu,
  extUriMatches,
  parseOnOff,
  stripColorHashPrefix,
  xsdDashStyle,
  xsdTextAnchor,
  xsdTextVerticalType,
} from "@office-open/core";
import type { UniversalMeasure } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { parse, stringify } from "@office-open/core/descriptor";
import type { ReadContext } from "@office-open/core/descriptor";
import {
  cell3DDesc,
  createBodyProperties,
  createTableStyle,
  fillDesc,
  findFillChild,
  parseTableStyle,
  outlineDesc,
  stringifyLineProperties,
  stringifyNonVisualDrawingProperties,
} from "@office-open/core/drawing";
import {
  attr,
  attrBool,
  attrMeasure,
  attrNum,
  children,
  escapeXml,
  findChild,
  textOf,
} from "@office-open/xml";
import type { Element } from "@office-open/xml";
import type {
  TableCellOptions,
  VerticalAlignment,
  TextVerticalType,
} from "@shared/table/table-cell";
import type { CellBorderOptions } from "@shared/table/table-cell-properties";
import type { TableOptions } from "@shared/table/table-frame";
import type { TableRowOptions } from "@shared/table/table-row";

import type { PptxWriteContext } from "../../context";
import {
  readGraphicFrameLocking,
  stringifyCnvGraphicFramePr,
  readNvPrPlaceholder,
  stringifyNvPr,
} from "./graphic-frame";
import { readCnvPr, readPositionFromXfrm } from "./shape";
import { paragraphDesc, type ParagraphDescriptorOptions } from "./text";

// ── Internal aliases ──

// The public types carry the canonical shapes; these name the inline
// border/margin maps the descriptor reads and writes cell-by-cell.
type CellBorders = NonNullable<TableCellOptions["borders"]>;
type CellMargins = NonNullable<TableCellOptions["margins"]>;
type TableLevelBorders = NonNullable<TableOptions["borders"]>;

// ── ID counter ──

let _nextTableId = 1024;

// ── Office 2014 table stamps (a16:colId on a:gridCol, a16:rowId on a:tr) ──

const A16_NS = "http://schemas.microsoft.com/office/drawing/2014/main";
const COL_ID_EXT_URI = "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}";
const ROW_ID_EXT_URI = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}";

type OfficeIdTag = "a16:colId" | "a16:rowId";

function officeIdExtLst(uri: string, tag: OfficeIdTag, val: string): string {
  return `<a:extLst><a:ext uri="${uri}"><${tag} xmlns:a16="${A16_NS}" val="${escapeXml(val)}"/></a:ext></a:extLst>`;
}

function readOfficeIdExt(el: Element, uri: string, tag: OfficeIdTag): string | undefined {
  const extLst = findChild(el, "a:extLst");
  if (!extLst) return undefined;
  for (const ext of extLst.elements ?? []) {
    if (ext.name !== "a:ext" || !extUriMatches(attr(ext, "uri"), uri)) continue;
    const idEl = findChild(ext, tag);
    const val = idEl ? attr(idEl, "val") : undefined;
    if (val) return val;
  }
  return undefined;
}

// ── Table (p:graphicFrame) descriptor ──

export const tableDesc: CustomDescriptor<TableOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const pptxCtx = ctx as PptxWriteContext;
    const id = opts.id ?? _nextTableId++;
    const name = opts.name ?? `Table ${id}`;

    const x = convertToEmu(opts.x ?? 0);
    const y = convertToEmu(opts.y ?? 0);
    // Default 100px when width/height unspecified
    const w = convertToEmu(opts.width ?? "100px");
    const h = convertToEmu(opts.height ?? "100px");

    const parts: string[] = [];

    // p:nvGraphicFramePr
    parts.push(
      `<p:nvGraphicFramePr>${stringifyNonVisualDrawingProperties("p:cNvPr", id, opts, name)}` +
        `${stringifyCnvGraphicFramePr(opts.locking)}` +
        `${stringifyNvPr(opts)}</p:nvGraphicFramePr>`,
    );

    // p:xfrm
    parts.push(`<p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></p:xfrm>`);

    // a:graphic > a:graphicData > a:tbl
    const tblParts: string[] = [];

    // a:tblPr
    tblParts.push(stringifyTblPr(opts));

    // a:tblGrid — gridCol widths are EMU; numbers stay (already EMU), strings convert.
    // Without explicit columnWidths the table fills its graphicFrame, so the
    // frame width is split evenly (a 0-width gridCol collapses the column).
    const colCount = opts.rows[0]?.cells.length ?? 1;
    const colWidths =
      opts.columnWidths && opts.columnWidths.length > 0
        ? opts.columnWidths.map(convertToEmu)
        : Array.from({ length: colCount }, () => Math.round(w / colCount));
    const gridCols = colWidths
      .map((cw, ci) => {
        const colId = opts.columnIds?.[ci];
        return colId
          ? `<a:gridCol w="${cw}">${officeIdExtLst(COL_ID_EXT_URI, "a16:colId", colId)}</a:gridCol>`
          : `<a:gridCol w="${cw}"/>`;
      })
      .join("");
    tblParts.push(`<a:tblGrid>${gridCols}</a:tblGrid>`);

    // a:tr[] — with border distribution
    const rowCount = opts.rows.length;
    for (const [ri, row] of opts.rows.entries()) {
      const cells = distributeBorders(row, ri, rowCount, opts.borders);
      tblParts.push(stringifyRow({ ...row, cells }, pptxCtx));
    }

    const tblXml = `<a:tbl>${tblParts.join("")}</a:tbl>`;
    parts.push(
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">${tblXml}</a:graphicData></a:graphic>`,
    );

    return `<p:graphicFrame>${parts.join("")}</p:graphicFrame>`;
  },

  parse(el, ctx) {
    const result: Partial<TableOptions> = {};

    // Position from p:xfrm
    const xfrm = findChild(el, "p:xfrm");
    if (xfrm) Object.assign(result, readPositionFromXfrm(xfrm));

    // Name from p:nvGraphicFramePr → p:cNvPr
    Object.assign(result, readCnvPr(el, "p:nvGraphicFramePr"));
    const locking = readGraphicFrameLocking(findChild(el, "p:nvGraphicFramePr"), ctx);
    readNvPrPlaceholder(findChild(el, "p:nvGraphicFramePr") ?? el, result);
    if (locking !== undefined) result.locking = locking;

    // Find a:tbl inside a:graphicData
    const graphicData = findChild(el, "a:graphic");
    const gd = graphicData ? findChild(graphicData, "a:graphicData") : undefined;
    const tbl = gd ? findChild(gd, "a:tbl") : undefined;
    if (!tbl) return result as TableOptions;

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
      } else {
        const tableStyleEl = findChild(tblPr, "a:tableStyle");
        if (tableStyleEl) {
          const style = parseTableStyle(tableStyleEl);
          if (style) result.tableStyle = style;
        }
      }

      // Table-level borders
      const borders: TableLevelBorders = {};
      for (const [elName, key] of [
        ["a:lnL", "left"],
        ["a:lnR", "right"],
        ["a:lnT", "top"],
        ["a:lnB", "bottom"],
      ] as const) {
        const borderEl = findChild(tblPr, elName);
        if (borderEl) {
          const borderOpts: Partial<CellBorderOptions> = {};
          const w = attrNum(borderEl, "w");
          if (w !== undefined) borderOpts.width = w;
          const fillResult = parseBorderColor(borderEl, ctx);
          if (fillResult !== undefined) borderOpts.color = fillResult;
          const prstDash = findChild(borderEl, "a:prstDash");
          if (prstDash) {
            const val = attr(prstDash, "val");
            if (val)
              borderOpts.dashStyle = xsdDashStyle.from(val) as CellBorderOptions["dashStyle"];
          }
          borderOpts.outline = parse(outlineDesc, borderEl, ctx);
          borders[key] = borderOpts as CellBorderOptions;
        }
      }
      if (Object.keys(borders).length > 0) result.borders = borders;
    }

    // a:tblGrid → columnWidths (+ per-column a16:colId stamps)
    const tblGrid = findChild(tbl, "a:tblGrid");
    if (tblGrid) {
      const colWidths: number[] = [];
      const columnIds: string[] = [];
      for (const gridCol of children(tblGrid, "a:gridCol")) {
        const w = attrNum(gridCol, "w");
        colWidths.push(w ?? 0);
        const colId = readOfficeIdExt(gridCol, COL_ID_EXT_URI, "a16:colId");
        columnIds.push(colId ?? "");
      }
      if (colWidths.length > 0) result.columnWidths = colWidths;
      if (columnIds.some((id) => id !== "")) result.columnIds = columnIds;
    }

    // a:tr → rows
    const rows: TableRowOptions[] = [];
    for (const tr of children(tbl, "a:tr")) {
      const h = attrNum(tr, "h");
      const cells: TableCellOptions[] = [];
      for (const tc of children(tr, "a:tc")) {
        cells.push(parseTableCell(tc, ctx));
      }
      const row: TableRowOptions = { cells };
      if (h !== undefined) row.height = h;
      const rowId = readOfficeIdExt(tr, ROW_ID_EXT_URI, "a16:rowId");
      if (rowId) row.rowId = rowId;
      rows.push(row);
    }
    result.rows = rows;

    return result as TableOptions;
  },
};

// ── Helpers ──

function stringifyTblPr(opts: TableOptions): string {
  const attrs: string[] = [];
  if (opts.firstRow !== undefined) attrs.push(`firstRow="${opts.firstRow ? 1 : 0}"`);
  if (opts.lastRow !== undefined) attrs.push(`lastRow="${opts.lastRow ? 1 : 0}"`);
  if (opts.bandRow !== undefined) attrs.push(`bandRow="${opts.bandRow ? 1 : 0}"`);
  if (opts.firstCol !== undefined) attrs.push(`firstCol="${opts.firstCol ? 1 : 0}"`);
  if (opts.lastCol !== undefined) attrs.push(`lastCol="${opts.lastCol ? 1 : 0}"`);
  if (opts.bandCol !== undefined) attrs.push(`bandCol="${opts.bandCol ? 1 : 0}"`);
  const styleId = opts.tableStyleId
    ? `<a:tableStyleId>${escapeXml(opts.tableStyleId)}</a:tableStyleId>`
    : "";
  // Inline CT_TableStyle — alternative to referencing a tableStyles part entry.
  const style = opts.tableStyle ? createTableStyle(opts.tableStyle, "a:tableStyle") : "";
  if (attrs.length === 0 && !styleId && !style) return "<a:tblPr/>";
  return `<a:tblPr ${attrs.join(" ")}>${styleId}${style}</a:tblPr>`;
}

function stringifyRow(row: TableRowOptions, ctx: PptxWriteContext): string {
  const h = convertToEmu(row.height ?? 0);
  const cellParts: string[] = [];
  for (const cell of row.cells) {
    cellParts.push(stringifyCell(cell, ctx));
  }
  // CT_TableRow tail — verbatim row stamp (a16:rowId's home).
  const ext = row.rowId ? officeIdExtLst(ROW_ID_EXT_URI, "a16:rowId", row.rowId) : "";
  return `<a:tr h="${h}">${cellParts.join("")}${ext}</a:tr>`;
}

function stringifyCell(cell: TableCellOptions, ctx: PptxWriteContext): string {
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

function stringifyTxBody(cell: TableCellOptions, ctx: PptxWriteContext): string {
  const txParts: string[] = [];

  txParts.push(createBodyProperties({}));
  txParts.push("<a:lstStyle/>");

  // Paragraphs — a:txBody requires at least one a:p, so an empty children
  // array (or paragraphs that stringify to nothing) falls back to a bare one
  const paragraphs: string[] = [];
  if (cell.children) {
    for (const c of cell.children) {
      const pXml =
        typeof c === "string"
          ? paragraphDesc.stringify({ children: [{ text: c }] }, ctx)
          : paragraphDesc.stringify(c, ctx);
      if (pXml) paragraphs.push(pXml);
    }
  } else if (cell.text !== undefined) {
    const pXml = paragraphDesc.stringify({ children: [{ text: cell.text }] }, ctx);
    if (pXml) paragraphs.push(pXml);
  }
  txParts.push(paragraphs.length > 0 ? paragraphs.join("") : "<a:p/>");

  return `<a:txBody>${txParts.join("")}</a:txBody>`;
}

function stringifyTcPr(cell: TableCellOptions, ctx: PptxWriteContext): string {
  const attrs: string[] = [];
  const children: string[] = [];

  if (cell.verticalAlign) attrs.push(`anchor="${xsdTextAnchor.to(cell.verticalAlign)}"`);
  if (cell.vertical) attrs.push(`vert="${xsdTextVerticalType.to(cell.vertical)}"`);
  if (cell.margins?.left !== undefined) attrs.push(`marL="${convertToEmu(cell.margins.left)}"`);
  if (cell.margins?.right !== undefined) attrs.push(`marR="${convertToEmu(cell.margins.right)}"`);
  if (cell.margins?.top !== undefined) attrs.push(`marT="${convertToEmu(cell.margins.top)}"`);
  if (cell.margins?.bottom !== undefined) attrs.push(`marB="${convertToEmu(cell.margins.bottom)}"`);

  if (cell.borders) {
    if (cell.borders.left) children.push(buildBorderLine("a:lnL", cell.borders.left, ctx));
    if (cell.borders.right) children.push(buildBorderLine("a:lnR", cell.borders.right, ctx));
    if (cell.borders.top) children.push(buildBorderLine("a:lnT", cell.borders.top, ctx));
    if (cell.borders.bottom) children.push(buildBorderLine("a:lnB", cell.borders.bottom, ctx));
    if (cell.borders.diagonalTopLeftToBottomRight)
      children.push(buildBorderLine("a:lnTlToBr", cell.borders.diagonalTopLeftToBottomRight, ctx));
    if (cell.borders.diagonalBottomLeftToTopRight)
      children.push(buildBorderLine("a:lnBlToTr", cell.borders.diagonalBottomLeftToTopRight, ctx));
  }

  if (cell.cell3D) {
    const cell3DXml = stringify(cell3DDesc, cell.cell3D, ctx);
    if (cell3DXml) children.push(cell3DXml);
  }

  if (cell.fill !== undefined) {
    const fillXml = stringify(fillDesc, cell.fill, ctx);
    if (fillXml) children.push(fillXml);
  }

  // A source cell may carry no a:tcPr at all (pandoc-style minimal output) —
  // only the round-trip marker re-emits the bare element.
  if (attrs.length === 0 && children.length === 0) return cell.cellProperties ? "<a:tcPr/>" : "";

  const attrStr = attrs.length > 0 ? ` ${attrs.join(" ")}` : "";
  if (children.length === 0) return `<a:tcPr${attrStr}/>`;
  return `<a:tcPr${attrStr}>${children.join("")}</a:tcPr>`;
}

function buildBorderLine(name: string, options: CellBorderOptions, ctx: PptxWriteContext): string {
  // Full line properties win when present (parsed sources keep every child).
  if (options.outline) {
    return stringifyLineProperties(name, options.outline, ctx) ?? `<${name}/>`;
  }
  const attrs: string[] = [];
  if (options.width !== undefined) attrs.push(`w="${convertToEmu(options.width)}"`);

  const children: string[] = [];
  if (options.color !== undefined) {
    const fillXml =
      typeof options.color === "string"
        ? `<a:solidFill><a:srgbClr val="${stripColorHashPrefix(options.color)}"/></a:solidFill>`
        : stringify(fillDesc, options.color, ctx);
    if (fillXml) children.push(fillXml);
  }
  if (options.dashStyle) {
    children.push(`<a:prstDash val="${xsdDashStyle.to(options.dashStyle)}"/>`);
  }

  const attrStr = attrs.length > 0 ? ` ${attrs.join(" ")}` : "";
  if (children.length === 0) return `<${name}${attrStr}/>`;
  return `<${name}${attrStr}>${children.join("")}</${name}>`;
}

/** Distribute table-level borders to edge cells. */
function distributeBorders(
  row: TableRowOptions,
  ri: number,
  rowCount: number,
  tb: TableOptions["borders"],
): TableCellOptions[] {
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
function parseTableCell(tc: Element, readCtx?: ReadContext): TableCellOptions {
  const result: TableCellOptions = {};
  const ctx = readCtx ?? ({} as ReadContext);

  const gridSpan = attrNum(tc, "gridSpan");
  if (gridSpan !== undefined && gridSpan > 1) result.columnSpan = gridSpan;
  const rowSpan = attrNum(tc, "rowSpan");
  if (rowSpan !== undefined && rowSpan > 1) result.rowSpan = rowSpan;
  const hMerge = parseOnOff(attr(tc, "hMerge"));
  if (hMerge === true) result.horizontalMerge = "restart";
  else if (hMerge === false) result.horizontalMerge = "continue";
  const vMerge = parseOnOff(attr(tc, "vMerge"));
  if (vMerge === true) result.verticalMerge = "restart";
  else if (vMerge === false) result.verticalMerge = "continue";

  // a:txBody → paragraph children
  const txBody = findChild(tc, "a:txBody");
  if (txBody) {
    const paragraphs: ParagraphDescriptorOptions[] = [];
    for (const pEl of txBody.elements ?? []) {
      if (pEl.name !== "a:p") continue;
      const para = paragraphDesc.parse(pEl, ctx);
      paragraphs.push(para);
    }

    if (paragraphs.length === 1) {
      const [p] = paragraphs;
      if (p) {
        if (p.text && !p.children) {
          result.text = p.text;
        } else {
          result.children = [p];
        }
      }
    } else if (paragraphs.length > 0) {
      result.children = paragraphs;
    }

    // Extract margins from a:bodyPr
    const bodyPr = findChild(txBody, "a:bodyPr");
    if (bodyPr) {
      const margins: CellMargins = {};
      const tIns = attrMeasure(bodyPr, "tIns");
      if (tIns !== undefined) margins.top = tIns as number | UniversalMeasure;
      const bIns = attrMeasure(bodyPr, "bIns");
      if (bIns !== undefined) margins.bottom = bIns as number | UniversalMeasure;
      const lIns = attrMeasure(bodyPr, "lIns");
      if (lIns !== undefined) margins.left = lIns as number | UniversalMeasure;
      const rIns = attrMeasure(bodyPr, "rIns");
      if (rIns !== undefined) margins.right = rIns as number | UniversalMeasure;
      if (Object.keys(margins).length > 0) result.margins = margins;
    }
  }

  // a:tcPr
  const tcPr = findChild(tc, "a:tcPr");
  if (tcPr) {
    const keysBefore = Object.keys(result).length;
    const anchor = attr(tcPr, "anchor");
    if (anchor) result.verticalAlign = xsdTextAnchor.from(anchor) as VerticalAlignment;
    const vert = attr(tcPr, "vert");
    if (vert) result.vertical = xsdTextVerticalType.from(vert) as TextVerticalType;

    // Margins from tcPr attributes
    const margins: CellMargins = {};
    const marL = attrMeasure(tcPr, "marL");
    if (marL !== undefined) margins.left = marL as number | UniversalMeasure;
    const marR = attrMeasure(tcPr, "marR");
    if (marR !== undefined) margins.right = marR as number | UniversalMeasure;
    const marT = attrMeasure(tcPr, "marT");
    if (marT !== undefined) margins.top = marT as number | UniversalMeasure;
    const marB = attrMeasure(tcPr, "marB");
    if (marB !== undefined) margins.bottom = marB as number | UniversalMeasure;
    if (Object.keys(margins).length > 0) result.margins = margins;

    // Fill — guard against fillDesc returning { type: "none" } for a tcPr with
    // no fill child, which would spuriously emit <a:noFill/> on re-stringify.
    const fillChild = findFillChild(tcPr);
    if (fillChild) result.fill = parse(fillDesc, fillChild, ctx);

    const cell3DEl = findChild(tcPr, "a:cell3D");
    if (cell3DEl) result.cell3D = parse(cell3DDesc, cell3DEl, ctx);

    // Cell borders
    const borders: CellBorders = {};
    for (const [elName, key] of [
      ["a:lnL", "left"],
      ["a:lnR", "right"],
      ["a:lnT", "top"],
      ["a:lnB", "bottom"],
      ["a:lnTlToBr", "diagonalTopLeftToBottomRight"],
      ["a:lnBlToTr", "diagonalBottomLeftToTopRight"],
    ] as const) {
      const borderEl = findChild(tcPr, elName);
      if (borderEl) {
        const borderOpts: Partial<CellBorderOptions> = {};
        const w = attrNum(borderEl, "w");
        if (w !== undefined) borderOpts.width = w;
        const fillResult = parseBorderColor(borderEl, ctx);
        if (fillResult !== undefined) borderOpts.color = fillResult;
        // Dash style
        const prstDash = findChild(borderEl, "a:prstDash");
        if (prstDash) {
          const val = attr(prstDash, "val");
          if (val) borderOpts.dashStyle = xsdDashStyle.from(val) as CellBorderOptions["dashStyle"];
        }
        // Full line properties — keeps joins/line ends and bare <a:lnX><a:noFill/></a:lnX>.
        borderOpts.outline = parse(outlineDesc, borderEl, ctx);
        borders[key] = borderOpts as CellBorderOptions;
      }
    }
    if (Object.keys(borders).length > 0) result.borders = borders;

    // A bare <a:tcPr/> yields no fields — mark the presence so stringify
    // re-emits the empty element.
    if (Object.keys(result).length === keysBefore) result.cellProperties = true;
  }

  return result;
}

/**
 * Read a border line color from its a:lnL/a:lnR/… element. An srgb solid fill
 * stays the hex-string sugar; any other fill (scheme color, gradient) comes
 * back as the structured FillOptions so it round-trips.
 */
function parseBorderColor(el: Element, ctx: ReadContext): CellBorderOptions["color"] {
  const solidFill = findChild(el, "a:solidFill");
  if (!solidFill) return undefined;
  const srgbClr = findChild(solidFill, "a:srgbClr");
  if (srgbClr?.attributes?.["val"] !== undefined) return String(srgbClr.attributes["val"]);
  const fill = parse(fillDesc, el, ctx);
  return fill ?? undefined;
}
