/**
 * XLSX parsing — parse .xlsx files into structured data.
 *
 * @module
 */
import { parseArchive, parseCorePropsElement } from "@office-open/core";
import type { ParsedArchive } from "@office-open/core";
import { toUint8Array } from "@office-open/core";
import type { DataType } from "@office-open/core";
import type { Element } from "@office-open/xml";
import { attr, findChild, textOf } from "@office-open/xml";
import { worksheetDesc } from "@parts/worksheet";
import type { WorksheetOptions } from "@parts/worksheet";
import type { WorkbookOptions } from "@shared/file";

import { XlsxReadContext } from "./context";

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

  // Create read context for descriptor pipeline
  const readContext = new XlsxReadContext(xlsx, strings);

  // Sheet names from workbook
  const sheetNames: string[] = [];
  if (xlsx.workbook) {
    const sheetsEl = findChild(xlsx.workbook, "sheets");
    if (sheetsEl) {
      for (const s of sheetsEl.elements ?? []) {
        if (s.name !== "sheet") continue;
        sheetNames.push(attr(s, "name") ?? "");
      }
    }
  }

  // Parse worksheets using descriptor pipeline
  const worksheets: WorksheetOptions[] = [];
  for (let i = 0; i < xlsx.worksheets.length; i++) {
    const wsPath = xlsx.worksheets[i];
    const wsEl = xlsx.doc.get(wsPath);
    if (!wsEl) continue;

    const wsOpts = worksheetDesc.parse(wsEl, readContext) as Record<string, unknown>;
    if (sheetNames[i]) wsOpts.name = sheetNames[i];
    worksheets.push(wsOpts as WorksheetOptions);
  }

  opts.worksheets = worksheets;
  return opts as WorkbookOptions;
}
