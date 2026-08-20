/**
 * Table (w:tbl) descriptor for DOCX.
 *
 * Stringifies pure JSON TableOptions into XML using direct string
 * concatenation — no intermediate object tree, no xml() pipeline.
 *
 * @module
 */

import { convertToTwip } from "@office-open/core";
import { xsdVerticalMergeRev } from "@office-open/core";
import type { PositiveUniversalMeasure, UniversalMeasure } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrBool, attrMeasure, attrNum, children, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import {
  parseCustomXmlProperties,
  stringifyCustomXmlShell,
  stringifySdtShell,
} from "@parts/bodychildren";
import type { CustomXmlCellOptions, CustomXmlRowOptions } from "@parts/custom-xml";
import {
  buildBookmarkStartAttrs,
  buildMarkupRangeAttrs,
  stringifyParagraphInline,
} from "@parts/inline";
import type { BookmarkStartOptions, MarkupRangeOptions } from "@parts/paragraph/links/bookmark";
import { parseRunProperties } from "@parts/paragraph/run/run-parse";
import { parseSdtProperties } from "@parts/sdt/sdt-parse";
import type { TableGridChangeOptions } from "@parts/table/grid";
import type { TableOptions } from "@parts/table/table";
import type { TableCellSpacingProperties } from "@parts/table/table-cell-spacing";
import type { SdtCellOptions, TableCellOptions } from "@parts/table/table-cell/table-cell";
import type { TableCellBordersOptions } from "@parts/table/table-cell/table-cell-components";
import { VerticalMergeType } from "@parts/table/table-cell/table-cell-components";
import type { TableBordersOptions } from "@parts/table/table-properties/table-borders";
import type { TableCellMarginOptions } from "@parts/table/table-properties/table-cell-margin";
import type { TableFloatOptions } from "@parts/table/table-properties/table-float-properties";
import type { TableLookOptions } from "@parts/table/table-properties/table-look";
import type {
  TablePropertyExChangeOptions,
  TablePropertyExOptions,
} from "@parts/table/table-properties/table-property-exceptions";
import type {
  SdtRowOptions,
  TableRowMarkerChild,
  TableRowOptions,
} from "@parts/table/table-row/table-row";
import type { CnfStyleOptions } from "@parts/table/table-row/table-row-properties";
import type { TableWidthProperties } from "@parts/table/table-width";
import { widthFiftiethsToPct } from "@parts/table/table-width";
import { parseBorderSide } from "@shared/border";
import type { SectionChild } from "@shared/section";
import { parseShading, type ShadingProperties } from "@shared/shading";
import type { CellMergeAttributes } from "@shared/track-revision";
import type { ChangedProperties } from "@shared/track-revision/track-revision";

import type { BodyContext, DocxReadContext } from "../../context";
import {
  stringifyTableCellProperties,
  stringifyTableProperties,
  stringifyTablePropertyExceptions,
  stringifyTableRowProperties,
} from "./stringify";
import type {
  TableCellPropertiesChangeOptions,
  TableCellPropertiesOptions,
} from "./table-cell/table-cell-properties";
import type {
  TablePropertiesChangeOptions,
  TablePropertiesOptions,
} from "./table-properties/table-properties";
import type {
  TableRowPropertiesChangeOptions,
  TableRowPropertiesOptions,
} from "./table-row/table-row-properties";

/** Parse track-change attributes (id/author/date) from w:ins/w:del/w:cellIns/w:cellDel. */
function parseChangeAttrs(el: Element): Partial<ChangedProperties> {
  const change: Partial<ChangedProperties> = {};
  const id = attrNum(el, "w:id");
  if (id !== undefined) change.id = id;
  const author = attr(el, "w:author");
  if (author) change.author = author;
  const date = attr(el, "w:date");
  if (date) change.date = date;
  return change;
}

// ── Table grid ──

function buildTableGridXml(
  widths: Array<number | string>,
  revision?: TableGridChangeOptions,
): string {
  // w:gridCol @w is s:ST_TwipsMeasure — normalize UniversalMeasure to twip integers
  // so the emitted @w is always a bare integer (never "25mm").
  const gridCol = (w: number | string): string =>
    `<w:gridCol w:w="${convertToTwip(w as number | UniversalMeasure)}"/>`;
  const cols = widths.map(gridCol).join("");

  if (revision) {
    const revCols = revision.columnWidths.map(gridCol).join("");
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
 * Handles strings and plain objects (paragraph/table); every other variant
 * (block sdt, customXml, altChunk, toc, textbox, rawXml) goes through the
 * body-child dispatcher the context carries — the same channel sdtBlockDesc
 * uses for its own children, avoiding a table ↔ body circular import.
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
  return ctx.stringifyChild(child);
}

// ── Cell stringification ──

function stringifyTableCell(cell: TableCellOptions, ctx: BodyContext): string {
  let body = "";

  const tcPr = stringifyTableCellProperties(cell);
  if (tcPr) body += tcPr;

  const children = cell.children as SectionChild[] | undefined;
  if (children) {
    for (const child of children) {
      body += stringifyCellChild(child, ctx);
    }
  }

  // Cells must end with a paragraph unless the last child is already one.
  // Block containers (sdt/customXml/toc/textbox) hold the closing paragraph
  // themselves, so a cell ending in one round-trips without an injected <w:p/>.
  const last = children?.[children.length - 1];
  const endsWithParagraph =
    last !== undefined &&
    typeof last !== "string" &&
    ("paragraph" in last ||
      "table" in last ||
      "sdt" in last ||
      "customXml" in last ||
      "toc" in last ||
      "textbox" in last);
  if (!endsWithParagraph) {
    body += "<w:p/>";
  }

  return `<w:tc>${body}</w:tc>`;
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

  // Cells (a cell may be wrapped by a cell-level SDT or customXml; run-level
  // markers anchored between cells emit at their captured position)
  for (const cell of row.cells) {
    if ("sdt" in cell) {
      const s = cell.sdt;
      const contentXml = (s.cells ?? []).map((c) => stringifyTableCell(c, ctx)).join("");
      parts.push(stringifySdtShell(s.properties, s.endProperties, contentXml));
    } else if ("customXml" in cell) {
      const cx = cell.customXml;
      const contentXml = (cx.children ?? []).map((c) => stringifyTableCell(c, ctx)).join("");
      parts.push(stringifyCustomXmlShell(cx, contentXml));
    } else if ("commentRangeStart" in cell) {
      parts.push(`<w:commentRangeStart ${buildMarkupRangeAttrs(cell.commentRangeStart)}/>`);
    } else if ("commentRangeEnd" in cell) {
      parts.push(`<w:commentRangeEnd ${buildMarkupRangeAttrs(cell.commentRangeEnd)}/>`);
    } else if ("bookmarkStart" in cell) {
      parts.push(`<w:bookmarkStart ${buildBookmarkStartAttrs(cell.bookmarkStart)}/>`);
    } else if ("bookmarkEnd" in cell) {
      parts.push(`<w:bookmarkEnd ${buildMarkupRangeAttrs(cell.bookmarkEnd)}/>`);
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
  let attr = "";
  if (row.runPropertiesRsid) attr += ` w:rsidRPr="${row.runPropertiesRsid}"`;
  if (row.additionRsid) attr += ` w:rsidR="${row.additionRsid}"`;
  if (row.deletionRsid) attr += ` w:rsidDel="${row.deletionRsid}"`;
  if (row.tableRowRsid) attr += ` w:rsidTr="${row.tableRowRsid}"`;

  const body = parts.join("");
  return body ? `<w:tr${attr}>${body}</w:tr>` : attr ? `<w:tr${attr}/>` : "<w:tr/>";
}

// ── Row options type ──

/** Type guard: a plain row (not SDT/customXml-wrapped). */
function isPlainRow(
  r: TableRowOptions | { sdt: SdtRowOptions } | { customXml: CustomXmlRowOptions },
): r is TableRowOptions {
  return !("sdt" in r) && !("customXml" in r);
}

/** Type guard: a run-level marker anchored between cells (not a cell). */
function isRowMarker(c: TableRowOptions["cells"][number]): c is TableRowMarkerChild {
  return (
    "commentRangeStart" in c || "commentRangeEnd" in c || "bookmarkStart" in c || "bookmarkEnd" in c
  );
}

/** Type guard: a plain cell (not SDT/customXml-wrapped, not a row marker). */
function isPlainCell(
  c:
    | TableCellOptions
    | { sdt: SdtCellOptions }
    | { customXml: CustomXmlCellOptions }
    | TableRowMarkerChild,
): c is TableCellOptions {
  if ("sdt" in c || "customXml" in c) return false;
  return !isRowMarker(c);
}

function findInsertIndex(
  cells: TableRowOptions["cells"],
  columnIndex: number,
  prefixCount: number,
): number {
  let colIdx = 0;
  for (const [i, c] of cells.entries()) {
    if (!isPlainCell(c)) continue; // SDT/customXml-wrapped cells don't occupy grid columns
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
  rows: TableOptions["rows"],
): Map<number, { cell: TableCellOptions; columnIndex: number }[]> {
  const extraMap = new Map<number, { cell: TableCellOptions; columnIndex: number }[]>();
  for (let ri = 0; ri < rows.length - 1; ri++) {
    const row = rows[ri];
    if (!row || !isPlainRow(row)) continue; // SDT/customXml-wrapped rows don't participate in merge
    const cells = row.cells;
    let colIdx = 0;

    for (const cell of cells) {
      if (!isPlainCell(cell)) continue; // SDT/customXml-wrapped cells don't participate in merge
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

/** Parse a w:tblCellMar / w:tcMar container into TableCellMarginOptions. */
function parseCellMargins(marginEl: Element): TableCellMarginOptions | undefined {
  const margins: TableCellMarginOptions = {};
  // CT_TblCellMar sides: top, start, left, bottom, end, right — each an
  // independent CT_TblWidth ({ size, type }). Parse faithfully: omit type
  // when the XML has none (stringify defaults it to DXA on the way back).
  for (const side of ["top", "start", "left", "bottom", "end", "right"] as const) {
    const sideEl = findChild(marginEl, `w:${side}`);
    if (sideEl) {
      const type = attr(sideEl, "w:type");
      const size = widthFiftiethsToPct(attrMeasure(sideEl, "w:w"), type);
      if (size !== undefined) {
        margins[side] = (
          type ? { size, type: type as TableWidthProperties["type"] } : { size }
        ) as TableWidthProperties;
      }
    }
  }
  if (Object.keys(margins).length === 0) return undefined;
  return margins as TableCellMarginOptions;
}

/** Parse a w:cnfStyle (CT_Cnf) element into CnfStyleOptions. */
function parseCnfStyle(cnfEl: Element): CnfStyleOptions | undefined {
  const cnf: CnfStyleOptions = {};
  const val = attr(cnfEl, "w:val");
  if (val) cnf.val = val;
  const firstRow = attrBool(cnfEl, "w:firstRow");
  if (firstRow !== undefined) cnf.firstRow = firstRow;
  const lastRow = attrBool(cnfEl, "w:lastRow");
  if (lastRow !== undefined) cnf.lastRow = lastRow;
  const firstColumn = attrBool(cnfEl, "w:firstColumn");
  if (firstColumn !== undefined) cnf.firstColumn = firstColumn;
  const lastColumn = attrBool(cnfEl, "w:lastColumn");
  if (lastColumn !== undefined) cnf.lastColumn = lastColumn;
  const oddVBand = attrBool(cnfEl, "w:oddVBand");
  if (oddVBand !== undefined) cnf.oddVBand = oddVBand;
  const evenVBand = attrBool(cnfEl, "w:evenVBand");
  if (evenVBand !== undefined) cnf.evenVBand = evenVBand;
  const oddHBand = attrBool(cnfEl, "w:oddHBand");
  if (oddHBand !== undefined) cnf.oddHBand = oddHBand;
  const evenHBand = attrBool(cnfEl, "w:evenHBand");
  if (evenHBand !== undefined) cnf.evenHBand = evenHBand;
  const firstRowFirstColumn = attrBool(cnfEl, "w:firstRowFirstColumn");
  if (firstRowFirstColumn !== undefined) cnf.firstRowFirstColumn = firstRowFirstColumn;
  const firstRowLastColumn = attrBool(cnfEl, "w:firstRowLastColumn");
  if (firstRowLastColumn !== undefined) cnf.firstRowLastColumn = firstRowLastColumn;
  const lastRowFirstColumn = attrBool(cnfEl, "w:lastRowFirstColumn");
  if (lastRowFirstColumn !== undefined) cnf.lastRowFirstColumn = lastRowFirstColumn;
  const lastRowLastColumn = attrBool(cnfEl, "w:lastRowLastColumn");
  if (lastRowLastColumn !== undefined) cnf.lastRowLastColumn = lastRowLastColumn;
  if (Object.keys(cnf).length === 0) return undefined;
  return cnf as CnfStyleOptions;
}

/**
 * Parse a w:tblPrEx (CT_TblPrEx) element into TablePropertyExOptions.
 * CT_TblPrExBase shares its child elements with CT_TblPrBase, so this reuses
 * parseTablePropertiesEl; both expose the margins field directly.
 */
function parseTablePropertyExceptions(el: Element): TablePropertyExOptions {
  const base = parseTablePropertiesEl(el);
  const opts: TablePropertyExOptions = {};
  if (base.width !== undefined) opts.width = base.width as TableWidthProperties;
  if (base.indent !== undefined) opts.indent = base.indent as TableWidthProperties;
  if (base.layout !== undefined) opts.layout = base.layout as TablePropertyExOptions["layout"];
  if (base.borders !== undefined) opts.borders = base.borders as TableBordersOptions;
  if (base.shading !== undefined) opts.shading = base.shading as ShadingProperties;
  if (base.alignment !== undefined) {
    opts.alignment = base.alignment as TablePropertyExOptions["alignment"];
  }
  if (base.margins !== undefined) opts.margins = base.margins;
  if (base.tableLook !== undefined) opts.tableLook = base.tableLook as TableLookOptions;
  if (base.cellSpacing !== undefined) {
    opts.cellSpacing = base.cellSpacing as TableCellSpacingProperties;
  }
  const tblPrExChange = findChild(el, "w:tblPrExChange");
  if (tblPrExChange) {
    const change = parseTablePropertyExChange(tblPrExChange);
    if (change) opts.tblPrExChange = change;
  }
  return opts as TablePropertyExOptions;
}

/** Parse a w:tblPrExChange (CT_TblPrExChange) — track-change wrapper around the previous tblPrEx. */
function parseTablePropertyExChange(el: Element): TablePropertyExChangeOptions | undefined {
  const change: Partial<TablePropertyExChangeOptions> = {};
  const id = attrNum(el, "w:id");
  if (id !== undefined) change.id = id;
  const author = attr(el, "w:author");
  if (author) change.author = author;
  const date = attr(el, "w:date");
  if (date) change.date = date;
  const innerTblPrEx = findChild(el, "w:tblPrEx");
  if (innerTblPrEx) {
    const inner = parseTablePropertyExceptions(innerTblPrEx);
    if (inner.width !== undefined) change.width = inner.width;
    if (inner.indent !== undefined) change.indent = inner.indent;
    if (inner.layout !== undefined) change.layout = inner.layout;
    if (inner.borders !== undefined) change.borders = inner.borders;
    if (inner.shading !== undefined) change.shading = inner.shading;
    if (inner.alignment !== undefined) change.alignment = inner.alignment;
    if (inner.margins !== undefined) change.margins = inner.margins;
    if (inner.tableLook !== undefined) change.tableLook = inner.tableLook;
    if (inner.cellSpacing !== undefined) change.cellSpacing = inner.cellSpacing;
  }
  if (change.id === undefined || change.author === undefined) return undefined;
  return change as TablePropertyExChangeOptions;
}

/** Map base 6-flags to TableLookOptions for w:tblLook emit — same field
 *  names and polarity, so this is a plain pass-through of the set flags. */
function tableLookFromFlags(opts: TableOptions): TableLookOptions | undefined {
  const has =
    opts.firstRow !== undefined ||
    opts.lastRow !== undefined ||
    opts.firstCol !== undefined ||
    opts.lastCol !== undefined ||
    opts.bandRow !== undefined ||
    opts.bandCol !== undefined;
  if (!has) return undefined;
  const look: TableLookOptions = {};
  if (opts.firstRow !== undefined) look.firstRow = opts.firstRow;
  if (opts.lastRow !== undefined) look.lastRow = opts.lastRow;
  if (opts.firstCol !== undefined) look.firstCol = opts.firstCol;
  if (opts.lastCol !== undefined) look.lastCol = opts.lastCol;
  if (opts.bandRow !== undefined) look.bandRow = opts.bandRow;
  if (opts.bandCol !== undefined) look.bandCol = opts.bandCol;
  return look;
}

/** Emit-side w:tblLook source: booleans may come from the base 6-flags
 *  (authoring shorthand) or the explicit tableLook (round-trip, which also
 *  carries the legacy w:val); merge both — undefined when neither is set. */
function tableLookForEmit(opts: TableOptions): TableLookOptions | undefined {
  const look: TableLookOptions = { ...opts.tableLook, ...tableLookFromFlags(opts) };
  return Object.keys(look).length > 0 ? look : undefined;
}

// ── Descriptor ──

export const tableDesc: CustomDescriptor<TableOptions, BodyContext> = {
  kind: "custom",

  stringify(opts, ctx) {
    const parts: string[] = [];

    // Table properties
    // tblPr is required in CT_Tbl (minOccurs defaults to 1) — always emit it,
    // even when empty; do not inject optional defaults (width/borders are XSD-optional).
    const tblPrOpts: TablePropertiesOptions & { includeIfEmpty?: boolean } = {
      alignment: opts.alignment,
      borders: opts.borders,
      caption: opts.caption,
      margins: opts.margins,
      cellSpacing: opts.cellSpacing,
      description: opts.description,
      float: opts.float,
      indent: opts.indent,
      layout: opts.layout,
      revision: opts.revision,
      shading: opts.shading,
      style: opts.style,
      styleColBandSize: opts.styleColBandSize,
      styleRowBandSize: opts.styleRowBandSize,
      tableLook: tableLookForEmit(opts),
      visuallyRightToLeft: opts.visuallyRightToLeft,
      width: opts.width,
      includeIfEmpty: true,
    };
    parts.push(stringifyTableProperties(tblPrOpts)!);

    // Table grid (fallback width count ignores markers between cells)
    const columnWidths =
      opts.columnWidths ??
      Array(
        Math.max(
          1,
          ...opts.rows.map((r) =>
            isPlainRow(r) ? r.cells.filter((c) => !isRowMarker(c)).length : 0,
          ),
        ),
      ).fill(100);
    parts.push(buildTableGridXml(columnWidths, opts.columnWidthsRevision));

    // Compute vertical merge CONTINUE cells
    const extraCells = computeVerticalMergeCells(opts.rows);

    // Rows (a row may be wrapped by a row-level SDT)
    for (const [ri, r] of opts.rows.entries()) {
      if ("sdt" in r) {
        const sdt = r.sdt;
        const contentXml = (sdt.rows ?? []).map((rr) => stringifyTableRow(rr, ctx)).join("");
        parts.push(stringifySdtShell(sdt.properties, sdt.endProperties, contentXml));
      } else if ("customXml" in r) {
        const cx = r.customXml;
        const contentXml = (cx.children ?? []).map((rr) => stringifyTableRow(rr, ctx)).join("");
        parts.push(stringifyCustomXmlShell(cx, contentXml));
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

export function parseTablePropertiesEl(el: Element): TablePropertiesOptions {
  const opts: TablePropertiesOptions = {};

  const style = findChild(el, "w:tblStyle");
  if (style) {
    const val = attr(style, "w:val");
    if (val) opts.style = val;
  }

  const tblW = findChild(el, "w:tblW");
  if (tblW) {
    const type = attr(tblW, "w:type");
    const size = widthFiftiethsToPct(attrMeasure(tblW, "w:w"), type);
    if (size !== undefined || type) {
      opts.width = { size: size ?? 0, ...(type ? { type } : {}) } as TableWidthProperties;
    }
  }

  const jc = findChild(el, "w:jc");
  if (jc) {
    const val = attr(jc, "w:val");
    if (val) opts.alignment = val as TablePropertiesOptions["alignment"];
  }

  const layout = findChild(el, "w:tblLayout");
  if (layout) {
    const val = attr(layout, "w:type");
    if (val === "autofit" || val === "fixed") opts.layout = val;
  }

  const tblBorders = findChild(el, "w:tblBorders");
  if (tblBorders) {
    // XML side names → TableBordersOptions keys (insideH/insideV map to
    // insideHorizontal/insideVertical); all 8 sides are CT_TblBorders-optional.
    const SIDE_KEYS: ReadonlyArray<[string, keyof TableBordersOptions]> = [
      ["top", "top"],
      ["start", "start"],
      ["left", "left"],
      ["bottom", "bottom"],
      ["end", "end"],
      ["right", "right"],
      ["insideH", "insideHorizontal"],
      ["insideV", "insideVertical"],
    ];
    const borders: TableBordersOptions = {};
    for (const [xmlSide, key] of SIDE_KEYS) {
      const sideEl = findChild(tblBorders, `w:${xmlSide}`);
      if (!sideEl) continue;
      const sideOpts = parseBorderSide(sideEl);
      if (!sideOpts) continue;
      borders[key] = sideOpts;
    }
    if (Object.keys(borders).length > 0) opts.borders = borders;
  }

  const tblCellMar = findChild(el, "w:tblCellMar");
  if (tblCellMar) {
    const margins = parseCellMargins(tblCellMar);
    if (margins) opts.margins = margins;
  }

  const shd = findChild(el, "w:shd");
  if (shd) {
    const shading = parseShading(shd);
    if (shading) opts.shading = shading;
  }

  // description → w:tblDescription/@w:val
  const tblDesc = findChild(el, "w:tblDescription");
  if (tblDesc) {
    const val = attr(tblDesc, "w:val");
    if (val) opts.description = val;
  }

  // float → w:tblpPr attributes + w:tblOverlap (sibling element in CT_TblPrBase)
  const tblpPr = findChild(el, "w:tblpPr");
  const tblOverlap = findChild(el, "w:tblOverlap");
  if (tblpPr || tblOverlap) {
    const floatOpts: Partial<TableFloatOptions> = {};
    if (tblpPr) {
      const horzAnchor = attr(tblpPr, "w:horzAnchor");
      if (horzAnchor)
        floatOpts.horizontalAnchor = horzAnchor as TableFloatOptions["horizontalAnchor"];
      const vertAnchor = attr(tblpPr, "w:vertAnchor");
      if (vertAnchor) floatOpts.verticalAnchor = vertAnchor as TableFloatOptions["verticalAnchor"];
      const tblpX = attrNum(tblpPr, "w:tblpX");
      if (tblpX !== undefined) floatOpts.absoluteHorizontalPosition = tblpX;
      const tblpXSpec = attr(tblpPr, "w:tblpXSpec");
      if (tblpXSpec)
        floatOpts.relativeHorizontalPosition =
          tblpXSpec as TableFloatOptions["relativeHorizontalPosition"];
      const tblpY = attrNum(tblpPr, "w:tblpY");
      if (tblpY !== undefined) floatOpts.absoluteVerticalPosition = tblpY;
      const tblpYSpec = attr(tblpPr, "w:tblpYSpec");
      if (tblpYSpec)
        floatOpts.relativeVerticalPosition =
          tblpYSpec as TableFloatOptions["relativeVerticalPosition"];
      const bottomFromText = attrNum(tblpPr, "w:bottomFromText");
      if (bottomFromText !== undefined) floatOpts.bottomFromText = bottomFromText;
      const topFromText = attrNum(tblpPr, "w:topFromText");
      if (topFromText !== undefined) floatOpts.topFromText = topFromText;
      const leftFromText = attrNum(tblpPr, "w:leftFromText");
      if (leftFromText !== undefined) floatOpts.leftFromText = leftFromText;
      const rightFromText = attrNum(tblpPr, "w:rightFromText");
      if (rightFromText !== undefined) floatOpts.rightFromText = rightFromText;
    }
    if (tblOverlap) {
      const overlap = attr(tblOverlap, "w:val");
      if (overlap) floatOpts.overlap = overlap as TableFloatOptions["overlap"];
    }
    if (Object.keys(floatOpts).length > 0) opts.float = floatOpts as TableFloatOptions;
  }

  // indent → w:tblInd/@w:w and @w:type
  const tblInd = findChild(el, "w:tblInd");
  if (tblInd) {
    const type = attr(tblInd, "w:type");
    const size = widthFiftiethsToPct(attrMeasure(tblInd, "w:w"), type);
    if (size !== undefined) {
      opts.indent = { size, ...(type ? { type } : {}) } as TableWidthProperties;
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
    const w = widthFiftiethsToPct(attrMeasure(tblCellSpacing, "w:w"), type);
    if (w !== undefined)
      opts.cellSpacing = { size: w, ...(type ? { type } : {}) } as TableCellSpacingProperties;
  }

  // Revision (w:tblPrChange)
  const tblPrChange = findChild(el, "w:tblPrChange");
  if (tblPrChange) {
    const rev: Partial<TablePropertiesChangeOptions> = {};
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
    if (Object.keys(rev).length > 0) opts.revision = rev as TablePropertiesChangeOptions;
  }

  // tblLook — conditional formatting flags (CT_TblLook). XML polarity
  // inverts banding (w:noHBand = !bandRow); w:val is the legacy hex
  // bitmask Word 2007 wrote alone.
  const tblLook = findChild(el, "w:tblLook");
  if (tblLook) {
    const look: TableLookOptions = {};
    const val = attr(tblLook, "w:val");
    if (val !== undefined) look.val = val;
    const firstRow = attrBool(tblLook, "w:firstRow");
    if (firstRow !== undefined) look.firstRow = firstRow;
    const lastRow = attrBool(tblLook, "w:lastRow");
    if (lastRow !== undefined) look.lastRow = lastRow;
    const firstColumn = attrBool(tblLook, "w:firstColumn");
    if (firstColumn !== undefined) look.firstCol = firstColumn;
    const lastColumn = attrBool(tblLook, "w:lastColumn");
    if (lastColumn !== undefined) look.lastCol = lastColumn;
    const noHBand = attrBool(tblLook, "w:noHBand");
    if (noHBand !== undefined) look.bandRow = !noHBand;
    const noVBand = attrBool(tblLook, "w:noVBand");
    if (noVBand !== undefined) look.bandCol = !noVBand;
    if (Object.keys(look).length > 0) opts.tableLook = look;
  }

  return opts;
}

function parseColumnWidthsEl(el: Element): {
  widths: Array<number | string>;
  revision?: TableGridChangeOptions;
} {
  const widths: Array<number | string> = [];
  const tblGrid = findChild(el, "w:tblGrid");
  if (!tblGrid) return { widths };

  for (const col of children(tblGrid, "w:gridCol")) {
    const w = attrMeasure(col, "w:w");
    widths.push(w ?? 100);
  }

  // CT_TblGrid may carry a tblGridChange (CT_TblGridChange = CT_Markup + tblGrid).
  const tblGridChange = findChild(tblGrid, "w:tblGridChange");
  if (tblGridChange) {
    const id = attrNum(tblGridChange, "w:id");
    const innerGrid = findChild(tblGridChange, "w:tblGrid");
    const revWidths: Array<number | string> = [];
    if (innerGrid) {
      for (const col of children(innerGrid, "w:gridCol")) {
        const w = attrMeasure(col, "w:w");
        revWidths.push(w ?? 100);
      }
    }
    if (id !== undefined) {
      return {
        widths,
        revision: { id, columnWidths: revWidths as number[] | PositiveUniversalMeasure[] },
      };
    }
  }

  return { widths };
}

export function parseTableRowPropertiesEl(el: Element): TableRowPropertiesOptions {
  const opts: TableRowPropertiesOptions = {};

  const trHeight = findChild(el, "w:trHeight");
  if (trHeight) {
    const val = attrMeasure(trHeight, "w:val");
    const rule = attr(trHeight, "w:hRule");
    if (val !== undefined) {
      opts.height = { value: val, ...(rule ? { rule } : {}) } as NonNullable<
        TableRowPropertiesOptions["height"]
      >;
    }
  }

  // cnfStyle → w:cnfStyle (CT_Cnf)
  const cnfStyle = findChild(el, "w:cnfStyle");
  if (cnfStyle) {
    const cnf = parseCnfStyle(cnfStyle);
    if (cnf) opts.cnfStyle = cnf;
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
    const type = attr(wBefore, "w:type");
    const size = widthFiftiethsToPct(attrMeasure(wBefore, "w:w"), type);
    if (size !== undefined)
      opts.widthBefore = { size, ...(type ? { type } : {}) } as TableWidthProperties;
  }
  const wAfter = findChild(el, "w:wAfter");
  if (wAfter) {
    const type = attr(wAfter, "w:type");
    const size = widthFiftiethsToPct(attrMeasure(wAfter, "w:w"), type);
    if (size !== undefined)
      opts.widthAfter = { size, ...(type ? { type } : {}) } as TableWidthProperties;
  }

  // rowAlignment → w:jc/@w:val
  const jc = findChild(el, "w:jc");
  if (jc) {
    const val = attr(jc, "w:val");
    if (val) opts.rowAlignment = val as TableRowPropertiesOptions["rowAlignment"];
  }

  // hidden → w:hidden
  const hidden = findChild(el, "w:hidden");
  if (hidden) opts.hidden = attrBool(hidden, "w:val") ?? true;

  // cellSpacing → w:tblCellSpacing
  const tblCellSpacing = findChild(el, "w:tblCellSpacing");
  if (tblCellSpacing) {
    const type = attr(tblCellSpacing, "w:type");
    const w = widthFiftiethsToPct(attrMeasure(tblCellSpacing, "w:w"), type);
    if (w !== undefined)
      opts.cellSpacing = { size: w, ...(type ? { type } : {}) } as TableCellSpacingProperties;
  }

  // insertion / deletion (track changes)
  const ins = findChild(el, "w:ins");
  if (ins) opts.insertion = parseChangeAttrs(ins) as ChangedProperties;
  const del = findChild(el, "w:del");
  if (del) opts.deletion = parseChangeAttrs(del) as ChangedProperties;

  // Revision (w:trPrChange)
  const trPrChange = findChild(el, "w:trPrChange");
  if (trPrChange) {
    const rev: Partial<TableRowPropertiesChangeOptions> = {};
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
    if (Object.keys(rev).length > 0) opts.revision = rev as TableRowPropertiesChangeOptions;
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

export function parseTableCellPropertiesEl(el: Element): TableCellPropertiesOptions {
  const opts: TableCellPropertiesOptions = {};

  const cnfStyle = findChild(el, "w:cnfStyle");
  if (cnfStyle) {
    const cnf = parseCnfStyle(cnfStyle);
    if (cnf) opts.cnfStyle = cnf;
  }

  const tcW = findChild(el, "w:tcW");
  if (tcW) {
    const type = attr(tcW, "w:type");
    const size = widthFiftiethsToPct(attrMeasure(tcW, "w:w"), type);
    if (size !== undefined) {
      opts.width = { size, ...(type ? { type } : {}) } as TableWidthProperties;
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
    if (val) opts.verticalAlign = val as TableCellPropertiesOptions["verticalAlign"];
  }

  const shd = findChild(el, "w:shd");
  if (shd) {
    const shading = parseShading(shd);
    if (shading) opts.shading = shading;
  }

  const tcBorders = findChild(el, "w:tcBorders");
  if (tcBorders) {
    // XML side name → TableCellBordersOptions key (incl. start/end + diagonals);
    // all sides are CT_TcBorders-optional. Mirrors table-level borders parse.
    const SIDE_KEYS: ReadonlyArray<[string, keyof TableCellBordersOptions]> = [
      ["top", "top"],
      ["start", "start"],
      ["left", "left"],
      ["bottom", "bottom"],
      ["end", "end"],
      ["right", "right"],
      ["insideH", "insideHorizontal"],
      ["insideV", "insideVertical"],
      ["tl2br", "topLeftToBottomRight"],
      ["tr2bl", "topRightToBottomLeft"],
    ];
    const borders: TableCellBordersOptions = {};
    for (const [xmlSide, key] of SIDE_KEYS) {
      const sideEl = findChild(tcBorders, `w:${xmlSide}`);
      if (!sideEl) continue;
      const sideOpts = parseBorderSide(sideEl);
      if (!sideOpts) continue;
      borders[key] = sideOpts;
    }
    if (Object.keys(borders).length > 0) opts.borders = borders;
  }

  const noWrap = findChild(el, "w:noWrap");
  if (noWrap) opts.noWrap = attrBool(noWrap, "w:val") ?? true;

  const tcMar = findChild(el, "w:tcMar");
  if (tcMar) {
    const margins = parseCellMargins(tcMar);
    if (margins) opts.margins = margins;
  }

  const textDirection = findChild(el, "w:textDirection");
  if (textDirection) {
    const val = attr(textDirection, "w:val");
    if (val) opts.textDirection = val as TableCellPropertiesOptions["textDirection"];
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
  if (cellIns) opts.insertion = parseChangeAttrs(cellIns) as ChangedProperties;
  const cellDel = findChild(el, "w:cellDel");
  if (cellDel) opts.deletion = parseChangeAttrs(cellDel) as ChangedProperties;

  // Revision (w:tcPrChange)
  const tcPrChange = findChild(el, "w:tcPrChange");
  if (tcPrChange) {
    const rev: Partial<TableCellPropertiesChangeOptions> = {};
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
    if (Object.keys(rev).length > 0) opts.revision = rev as TableCellPropertiesChangeOptions;
  }

  // cellMerge → w:cellMerge
  const cellMerge = findChild(el, "w:cellMerge");
  if (cellMerge) {
    const cm = parseChangeAttrs(cellMerge) as Partial<CellMergeAttributes>;
    const vMerge = attr(cellMerge, "w:vMerge");
    if (vMerge) cm.verticalMerge = xsdVerticalMergeRev.from(vMerge) as "continue" | "restart";
    const vMergeOrig = attr(cellMerge, "w:vMergeOrig");
    if (vMergeOrig) {
      cm.verticalMergeOriginal = xsdVerticalMergeRev.from(vMergeOrig) as "continue" | "restart";
    }
    if (Object.keys(cm).length > 0) opts.cellMerge = cm as CellMergeAttributes;
  }

  return opts;
}

function parseTableCellEl(el: Element, ctx: DocxReadContext): TableCellOptions {
  const opts: Partial<TableCellOptions> = {};

  const tcPr = findChild(el, "w:tcPr");
  if (tcPr) {
    const props = parseTableCellPropertiesEl(tcPr);
    Object.assign(opts, props);
    // A bare <w:tcPr/> yields no fields — mark the presence so stringify
    // re-emits the empty element.
    if (Object.keys(props).length === 0) opts.cellProperties = true;
  }

  const childElements: SectionChild[] = [];
  for (const child of el.elements ?? []) {
    // Cell content is EG_BlockLevelElts — the same group the body holds, so
    // every non-property child (w:p, w:tbl, block sdt, customXml, altChunk,
    // subDoc, run-level markers) goes through the shared section-child parser.
    // Whitespace text nodes survive captureSpacesBetweenElements parsing and
    // must not reach the dispatcher (they stringify as empty rawXml noise).
    if (child.type !== "element" || child.name === "w:tcPr") continue;
    if (_parseChild) childElements.push(_parseChild(child, ctx));
  }

  opts.children = childElements;
  return opts as TableCellOptions;
}

function parseTableRowEl(el: Element, ctx: DocxReadContext): TableRowOptions {
  const opts: Partial<TableRowOptions> = {};

  const trPr = findChild(el, "w:trPr");
  if (trPr) {
    Object.assign(opts, parseTableRowPropertiesEl(trPr));
  }

  // w:tblPrEx (CT_TblPrEx) — per-row table-property exceptions
  const tblPrEx = findChild(el, "w:tblPrEx");
  if (tblPrEx) {
    const exceptions = parseTablePropertyExceptions(tblPrEx);
    if (Object.keys(exceptions).length > 0) opts.propertyExceptions = exceptions;
  }

  // rsid attributes on w:tr element
  for (const [attrName, optKey] of [
    ["w:rsidRPr", "runPropertiesRsid"],
    ["w:rsidR", "additionRsid"],
    ["w:rsidDel", "deletionRsid"],
    ["w:rsidTr", "tableRowRsid"],
  ] as const) {
    const val = attr(el, attrName);
    if (val) opts[optKey] = val;
  }

  const childCells: (
    | TableCellOptions
    | { sdt: SdtCellOptions }
    | { customXml: CustomXmlCellOptions }
    | TableRowMarkerChild
  )[] = [];
  for (const child of el.elements ?? []) {
    if (child.name === "w:tc") {
      childCells.push(parseTableCellEl(child, ctx));
    } else if (child.name === "w:commentRangeStart" || child.name === "w:commentRangeEnd") {
      // Markers anchored between cells (a comment/bookmark range spanning
      // cell boundaries) keep their position inside the cells sequence.
      const id = attrNum(child, "w:id");
      if (id !== undefined) {
        const marker: MarkupRangeOptions = { id };
        const disp = attr(child, "w:displacedByCustomXml");
        if (disp === "before" || disp === "after") marker.displacedByCustomXml = disp;
        childCells.push(
          child.name === "w:commentRangeStart"
            ? { commentRangeStart: marker }
            : { commentRangeEnd: marker },
        );
      }
    } else if (child.name === "w:bookmarkStart") {
      const id = attrNum(child, "w:id");
      const name = attr(child, "w:name");
      if (id !== undefined && name) {
        const bookmarkStart: Partial<BookmarkStartOptions> = { id, name };
        const disp = attr(child, "w:displacedByCustomXml");
        if (disp === "before" || disp === "after") bookmarkStart.displacedByCustomXml = disp;
        const colFirst = attrNum(child, "w:colFirst");
        if (colFirst !== undefined) bookmarkStart.colFirst = colFirst;
        const colLast = attrNum(child, "w:colLast");
        if (colLast !== undefined) bookmarkStart.colLast = colLast;
        childCells.push({ bookmarkStart: bookmarkStart as BookmarkStartOptions });
      }
    } else if (child.name === "w:bookmarkEnd") {
      const id = attrNum(child, "w:id");
      if (id !== undefined) {
        const bookmarkEnd: Partial<MarkupRangeOptions> = { id };
        const disp = attr(child, "w:displacedByCustomXml");
        if (disp === "before" || disp === "after") bookmarkEnd.displacedByCustomXml = disp;
        childCells.push({ bookmarkEnd: bookmarkEnd as MarkupRangeOptions });
      }
    } else if (child.name === "w:sdt") {
      const sdtPr = findChild(child, "w:sdtPr");
      const properties = sdtPr ? parseSdtProperties(sdtPr) : {};
      // CT_SdtEndPr wraps its run properties in a w:rPr child
      const sdtEndPr = findChild(child, "w:sdtEndPr");
      const endRPr = sdtEndPr ? findChild(sdtEndPr, "w:rPr") : undefined;
      const endProperties = sdtEndPr ? (endRPr ? parseRunProperties(endRPr) : {}) : undefined;
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
    } else if (child.name === "w:customXml") {
      const element = attr(child, "w:element") ?? "";
      const cx: CustomXmlCellOptions = { element };
      const cxUri = attr(child, "w:uri");
      if (cxUri) cx.uri = cxUri;
      const xmlPr = findChild(child, "w:customXmlPr");
      if (xmlPr) {
        const parsed = parseCustomXmlProperties(xmlPr);
        if (parsed.placeholder !== undefined || parsed.attributes !== undefined)
          cx.customXmlPr = parsed;
      }
      const cxCells: TableCellOptions[] = [];
      for (const sub of child.elements ?? []) {
        if (sub.name === "w:tc") cxCells.push(parseTableCellEl(sub, ctx));
      }
      if (cxCells.length > 0) cx.children = cxCells;
      childCells.push({ customXml: cx });
    }
  }

  opts.cells = childCells;
  return opts as TableRowOptions;
}

function parseTableEl(el: Element, ctx: DocxReadContext): TableOptions {
  const opts: Partial<TableOptions> = {};

  const tblPr = findChild(el, "w:tblPr");
  if (tblPr) {
    const tblPrParsed = parseTablePropertiesEl(tblPr);
    // tableLook carries the same 6-flag field names as TableOptions, so only
    // the emit-side (stringify) translates; here it assigns straight through.
    // The legacy w:val has no base-flag counterpart and stays on tableLook.
    const { tableLook, ...rest } = tblPrParsed;
    Object.assign(opts, rest);
    if (tableLook) {
      if (tableLook.firstRow !== undefined) opts.firstRow = tableLook.firstRow;
      if (tableLook.lastRow !== undefined) opts.lastRow = tableLook.lastRow;
      if (tableLook.firstCol !== undefined) opts.firstCol = tableLook.firstCol;
      if (tableLook.lastCol !== undefined) opts.lastCol = tableLook.lastCol;
      if (tableLook.bandRow !== undefined) opts.bandRow = tableLook.bandRow;
      if (tableLook.bandCol !== undefined) opts.bandCol = tableLook.bandCol;
      if (tableLook.val !== undefined) opts.tableLook = { val: tableLook.val };
    }
  }

  const grid = parseColumnWidthsEl(el);
  if (grid.widths.length > 0) {
    opts.columnWidths = grid.widths as NonNullable<TableOptions["columnWidths"]>;
  }
  if (grid.revision) {
    opts.columnWidthsRevision = grid.revision;
  }

  const rows: (TableRowOptions | { sdt: SdtRowOptions } | { customXml: CustomXmlRowOptions })[] =
    [];
  for (const child of el.elements ?? []) {
    if (child.name === "w:tr") {
      rows.push(parseTableRowEl(child, ctx));
    } else if (child.name === "w:sdt") {
      const sdtPr = findChild(child, "w:sdtPr");
      const properties = sdtPr ? parseSdtProperties(sdtPr) : {};
      // CT_SdtEndPr wraps its run properties in a w:rPr child
      const sdtEndPr = findChild(child, "w:sdtEndPr");
      const endRPr = sdtEndPr ? findChild(sdtEndPr, "w:rPr") : undefined;
      const endProperties = sdtEndPr ? (endRPr ? parseRunProperties(endRPr) : {}) : undefined;
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
    } else if (child.name === "w:customXml") {
      const element = attr(child, "w:element") ?? "";
      const cx: CustomXmlRowOptions = { element };
      const cxUri = attr(child, "w:uri");
      if (cxUri) cx.uri = cxUri;
      const xmlPr = findChild(child, "w:customXmlPr");
      if (xmlPr) {
        const parsed = parseCustomXmlProperties(xmlPr);
        if (parsed.placeholder !== undefined || parsed.attributes !== undefined)
          cx.customXmlPr = parsed;
      }
      const cxRows: TableRowOptions[] = [];
      for (const sub of child.elements ?? []) {
        if (sub.name === "w:tr") cxRows.push(parseTableRowEl(sub, ctx));
      }
      if (cxRows.length > 0) cx.children = cxRows;
      rows.push({ customXml: cx });
    }
  }

  opts.rows = rows;
  return opts as TableOptions;
}
