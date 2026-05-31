/**
 * XLSX parsing — parse .xlsx files into structured data.
 *
 * @module
 */
import { parseArchive, parseCorePropsElement } from "@office-open/core";
import type { ParsedArchive } from "@office-open/core";
import type { Element } from "@office-open/xml";
import { attr, attrNum, findChild, textOf } from "@office-open/xml";
import { toUint8Array } from "undio";
import type { DataType } from "undio";

import type { WorkbookOptions } from "./file/file";
import type {
  WorksheetOptions,
  RowOptions,
  CellOptions,
  ColumnOptions,
  MergeCellOptions,
} from "./file/worksheet";
import { letterToColumn } from "./util";

export { parseArchive };

// ── Low-level parse result ──

export interface XlsxPartRefs {
  worksheets: string[];
  charts: string[];
  media: string[];
  drawings: string[];
}

export interface XlsxDocument {
  doc: ParsedArchive;
  /** xl/workbook.xml root element */
  workbook?: Element;
  /** Worksheet paths (xl/worksheets/sheet{n}.xml) */
  worksheets: string[];
  /** xl/styles.xml root element */
  styles?: Element;
  /** xl/sharedStrings.xml root element */
  sharedStrings?: Element;
  partRefs: XlsxPartRefs;
  /** docProps/core.xml path */
  coreProps?: string;
  /** docProps/app.xml path */
  appProps?: string;
}

function sortByNumber(paths: string[]): string[] {
  return paths.sort((a, b) => {
    const numA = parseInt(a.match(/(\d+)/)?.[1] ?? "0", 10);
    const numB = parseInt(b.match(/(\d+)/)?.[1] ?? "0", 10);
    return numA - numB;
  });
}

/**
 * Parse raw .xlsx data into a low-level XlsxDocument.
 */
export function parseXlsx(data: DataType): XlsxDocument {
  const uint8 = toUint8Array(data);
  const doc = parseArchive(uint8);

  const workbook = doc.get("xl/workbook.xml");
  const styles = doc.get("xl/styles.xml");
  const sharedStrings = doc.get("xl/sharedStrings.xml");

  // Resolve worksheet paths from workbook rels
  const worksheets: string[] = [];
  const charts: string[] = [];
  const drawings: string[] = [];
  const media: string[] = [];

  const wbRels = doc.get("xl/_rels/workbook.xml.rels");
  if (wbRels) {
    for (const child of wbRels.elements ?? []) {
      if (child.name !== "Relationship") continue;
      const type = attr(child, "Type") ?? "";
      const target = attr(child, "Target") ?? "";
      if (!target) continue;

      if (type.includes("/worksheet")) {
        worksheets.push(target.startsWith("/") ? target.slice(1) : `xl/${target}`);
      }
    }
  }
  sortByNumber(worksheets);

  // Scan for drawings, charts, media
  drawings.push(...doc.keys("xl/drawings/").filter((k) => k.endsWith(".xml")));
  charts.push(...doc.keys("xl/charts/").filter((k) => k.endsWith(".xml")));
  media.push(...doc.keys("xl/media/"));
  sortByNumber(drawings);
  sortByNumber(charts);

  // Root rels → core/app props
  let coreProps: string | undefined;
  let appProps: string | undefined;
  const rootRels = doc.get("_rels/.rels");
  if (rootRels) {
    for (const child of rootRels.elements ?? []) {
      if (child.name !== "Relationship") continue;
      const type = attr(child, "Type") ?? "";
      const target = attr(child, "Target") ?? "";
      if (type.includes("/core-properties")) coreProps = target;
      else if (type.includes("/extended-properties")) appProps = target;
    }
  }

  return {
    doc,
    workbook,
    worksheets,
    styles,
    sharedStrings,
    partRefs: { worksheets, charts, media, drawings },
    coreProps,
    appProps,
  };
}

// ── Shared strings helper ──

function parseSharedStrings(el: Element | undefined): string[] {
  if (!el) return [];
  const strings: string[] = [];
  for (const si of el.elements ?? []) {
    if (si.name !== "si") continue;
    // Handle both <si><t>text</t></si> and <si><r><t>text</t></r>...</si>
    const t = findChild(si, "t");
    if (t) {
      strings.push(textOf(t) ?? "");
    } else {
      // Rich text: concatenate all <r><t> segments
      const parts: string[] = [];
      for (const r of si.elements ?? []) {
        if (r.name !== "r") continue;
        const rt = findChild(r, "t");
        if (rt) parts.push(textOf(rt) ?? "");
      }
      strings.push(parts.join(""));
    }
  }
  return strings;
}

// ── Cell reference helpers ──

function colFromRef(ref: string): number {
  const match = ref.match(/^([A-Z]+)/);
  return match ? letterToColumn(match[1]) : 1;
}

function rowFromRef(ref: string): number {
  const match = ref.match(/(\d+)$/);
  return match ? parseInt(match[1], 10) : 1;
}

// ── Worksheet parsing ──

function parseWorksheetElement(wsEl: Element, strings: string[]): WorksheetOptions {
  const opts: Record<string, unknown> = {};

  // Sheet name from the parent workbook's sheet element (not available here)
  // Caller should set it

  // Columns
  // Elements might be prefixed or unprefixed depending on how the XML parser handles default namespaces.
  const colsEl = findChild(wsEl, "cols") ?? findChildByLocalName(wsEl, "cols");
  if (colsEl) {
    const columns: ColumnOptions[] = [];
    for (const col of colsEl.elements ?? []) {
      if (localName(col) !== "col") continue;
      const min = attrNum(col, "min");
      const max = attrNum(col, "max");
      if (min === undefined || max === undefined) continue;
      const colOpts: Record<string, unknown> = { min, max };
      const width = attrNum(col, "width");
      if (width !== undefined) colOpts.width = width;
      if (attr(col, "hidden") === "1") colOpts.hidden = true;
      columns.push(colOpts as unknown as ColumnOptions);
    }
    if (columns.length > 0) opts.columns = columns;
  }

  // Freeze panes
  const sheetViews = findChildByLocalName(wsEl, "sheetViews");
  if (sheetViews) {
    const sheetView = findChildByLocalName(sheetViews, "sheetView");
    if (sheetView) {
      const pane = findChildByLocalName(sheetView, "pane");
      if (pane) {
        const state = attr(pane, "state");
        if (state === "frozen") {
          const freezePanes: Record<string, unknown> = {};
          const ySplit = attrNum(pane, "ySplit");
          const xSplit = attrNum(pane, "xSplit");
          if (ySplit && ySplit > 0) freezePanes.row = ySplit;
          if (xSplit && xSplit > 0) freezePanes.col = xSplit;
          if (Object.keys(freezePanes).length > 0) opts.freezePanes = freezePanes;
        }
      }
    }
  }

  // Sheet data (rows)
  const sheetData = findChildByLocalName(wsEl, "sheetData");
  const rows: RowOptions[] = [];
  if (sheetData) {
    for (const rowEl of sheetData.elements ?? []) {
      if (localName(rowEl) !== "row") continue;
      const rowNumber = attrNum(rowEl, "r");
      const rowOpts: Record<string, unknown> = {};
      if (rowNumber !== undefined) rowOpts.rowNumber = rowNumber;

      const ht = attrNum(rowEl, "ht");
      if (ht !== undefined) rowOpts.height = ht;
      if (attr(rowEl, "hidden") === "1") rowOpts.hidden = true;

      const cells: CellOptions[] = [];
      for (const cellEl of rowEl.elements ?? []) {
        if (localName(cellEl) !== "c") continue;
        const ref = attr(cellEl, "r");
        const type = attr(cellEl, "t");
        const cellOpts: Record<string, unknown> = {};
        if (ref) cellOpts.reference = ref;

        const styleIdx = attrNum(cellEl, "s");
        if (styleIdx !== undefined) cellOpts.styleIndex = styleIdx;

        // Cell value
        const vEl = findChildByLocalName(cellEl, "v");
        const isEl = findChildByLocalName(cellEl, "is");

        if (type === "s" && vEl) {
          // Shared string
          const idx = parseInt(textOf(vEl) ?? "", 10);
          cellOpts.value = strings[idx] ?? "";
        } else if (type === "b" && vEl) {
          cellOpts.value = textOf(vEl) === "1";
        } else if (type === "inlineStr" && isEl) {
          const t = findChildByLocalName(isEl, "t");
          cellOpts.value = textOf(t) ?? "";
        } else if (vEl) {
          const raw = textOf(vEl) ?? "";
          const num = Number(raw);
          cellOpts.value = isNaN(num) ? raw : num;
        }

        cells.push(cellOpts as CellOptions);
      }

      rowOpts.cells = cells;
      rows.push(rowOpts as RowOptions);
    }
  }
  opts.rows = rows;

  // Merge cells
  const mergeCellsEl = findChildByLocalName(wsEl, "mergeCells");
  if (mergeCellsEl) {
    const mergeCells: MergeCellOptions[] = [];
    for (const mc of mergeCellsEl.elements ?? []) {
      if (localName(mc) !== "mergeCell") continue;
      const ref = attr(mc, "ref");
      if (!ref) continue;
      const parts = ref.split(":");
      if (parts.length === 2) {
        mergeCells.push({
          from: { row: rowFromRef(parts[0]), col: colFromRef(parts[0]) },
          to: { row: rowFromRef(parts[1]), col: colFromRef(parts[1]) },
        });
      }
    }
    if (mergeCells.length > 0) opts.mergeCells = mergeCells;
  }

  // Auto filter
  const autoFilterEl = findChildByLocalName(wsEl, "autoFilter");
  if (autoFilterEl) {
    const ref = attr(autoFilterEl, "ref");
    if (ref) opts.autoFilter = ref;
  }

  return opts as WorksheetOptions;
}

// ── Helpers ──

function localName(el: Element): string {
  const name = el.name ?? "";
  const colonIdx = name.indexOf(":");
  return colonIdx >= 0 ? name.slice(colonIdx + 1) : name;
}

function findChildByLocalName(parent: Element, name: string): Element | undefined {
  return (parent.elements ?? []).find((el) => localName(el) === name);
}

/**
 * Parse a .xlsx file and convert it into WorkbookOptions.
 *
 * The returned options can be passed to `new Workbook(parsed)`.
 */
export function parseWorkbook(data: DataType): WorkbookOptions {
  const xlsx = parseXlsx(data);

  const opts: Record<string, unknown> = {};

  // Core properties
  if (xlsx.coreProps) {
    const corePropsEl = xlsx.doc.get(xlsx.coreProps);
    if (corePropsEl) {
      const cp = parseCorePropsElement(corePropsEl);
      if (cp.title) opts.title = cp.title;
      if (cp.subject) opts.subject = cp.subject;
      if (cp.creator) opts.creator = cp.creator;
      if (cp.keywords) opts.keywords = cp.keywords;
      if (cp.description) opts.description = cp.description;
      if (cp.lastModifiedBy) opts.lastModifiedBy = cp.lastModifiedBy;
      if (cp.revision) opts.revision = parseInt(cp.revision, 10);
    }
  }

  // Shared strings
  const strings = parseSharedStrings(xlsx.sharedStrings);

  // Sheet names from workbook
  const sheetNames: string[] = [];
  if (xlsx.workbook) {
    const sheetsEl = findChildByLocalName(xlsx.workbook, "sheets");
    if (sheetsEl) {
      for (const s of sheetsEl.elements ?? []) {
        if (localName(s) !== "sheet") continue;
        sheetNames.push(attr(s, "name") ?? "");
      }
    }
  }

  // Parse worksheets
  const worksheets: WorksheetOptions[] = [];
  for (let i = 0; i < xlsx.worksheets.length; i++) {
    const wsPath = xlsx.worksheets[i];
    const wsEl = xlsx.doc.get(wsPath);
    if (!wsEl) continue;

    const wsOpts = parseWorksheetElement(wsEl, strings) as Record<string, unknown>;
    if (sheetNames[i]) wsOpts.name = sheetNames[i];
    worksheets.push(wsOpts as WorksheetOptions);
  }

  opts.worksheets = worksheets;
  return opts as WorkbookOptions;
}
