/**
 * XLSX parsing — parse .xlsx files into structured data.
 *
 * @module
 */
import {
  appPropertiesDesc,
  customPropertiesDesc,
  parseArchive,
  parseCorePropsElement,
} from "@office-open/core";
import type { ParsedArchive } from "@office-open/core";
import { toUint8Array } from "@office-open/core";
import type { DataType } from "@office-open/core";
import type { ReadContext } from "@office-open/core/descriptor";
import type { Element } from "@office-open/xml";
import { attr } from "@office-open/xml";
import { calcChainDesc } from "@parts/calc-chain";
import { chartsheetDesc } from "@parts/chartsheet";
import type { ChartsheetOptions } from "@parts/chartsheet";
import { commentsDesc } from "@parts/comments";
import { drawingDesc } from "@parts/drawing";
import { externalLinkDesc } from "@parts/external-link";
import type { SharedWorkbookOptions, WorkbookOptions } from "@parts/file";
import { pivotCacheDefDesc, pivotCacheRecordsDesc } from "@parts/pivot-cache";
import { pivotTableDesc } from "@parts/pivot-table";
import {
  revisionHeadersDesc,
  revisionLogDesc,
  usersDesc,
  type RevisionLogOptions,
} from "@parts/revision-log";
import { sharedStringsDesc } from "@parts/shared-strings";
import { stylesDesc } from "@parts/styles";
import { tableDesc } from "@parts/table";
import { workbookDesc } from "@parts/workbook";
import { worksheetDesc } from "@parts/worksheet";
import type { WorksheetOptions } from "@parts/worksheet";

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
  /** docProps/custom.xml path */
  customProps?: string;
}

function sortByNumber(paths: string[]): string[] {
  return paths.sort((a, b) => {
    const numA = parseInt(a.match(/(\d+)/)?.[1] ?? "0", 10);
    const numB = parseInt(b.match(/(\d+)/)?.[1] ?? "0", 10);
    return numA - numB;
  });
}

/** Derive the sibling rels path for an external link part.
 * "xl/externalLinks/externalLink1.xml" → "xl/externalLinks/_rels/externalLink1.xml.rels"
 */
function externalLinkRelsPath(elPath: string): string {
  const idx = elPath.lastIndexOf("/");
  const dir = elPath.substring(0, idx);
  const file = elPath.substring(idx + 1);
  return `${dir}/_rels/${file}.rels`;
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
  let customProps: string | undefined;
  const rootRels = doc.get("_rels/.rels");
  if (rootRels) {
    for (const child of rootRels.elements ?? []) {
      if (child.name !== "Relationship") continue;
      const type = attr(child, "Type") ?? "";
      const target = attr(child, "Target") ?? "";
      if (type.includes("/core-properties")) coreProps = target;
      else if (type.includes("/extended-properties")) appProps = target;
      else if (type.includes("/custom-properties")) customProps = target;
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
    customProps,
  };
}

// ── Shared strings helper ──

/** Extract plain string array from sharedStringsDesc.parse() result. */
function extractStringsFromEntries(parsed: {
  entries: (string | Record<string, unknown>)[];
}): string[] {
  const strings: string[] = [];
  for (const entry of parsed.entries) {
    if (typeof entry === "string") {
      strings.push(entry);
    } else {
      // Rich text: concatenate run texts
      const runs = entry.runs as { text: string }[] | undefined;
      if (runs && runs.length > 0) {
        strings.push(runs.map((r) => r.text).join(""));
      }
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
      if (cp.revision !== undefined) opts.revision = cp.revision;
      if (cp.lastPrinted) opts.lastPrinted = cp.lastPrinted;
      if (cp.created) opts.created = cp.created;
      if (cp.modified) opts.modified = cp.modified;
    }
  }

  // Extended (app) properties
  if (xlsx.appProps) {
    const appPropsEl = xlsx.doc.get(xlsx.appProps);
    if (appPropsEl) {
      const ap = appPropertiesDesc.parse(appPropsEl, {} as ReadContext);
      if (ap && Object.keys(ap).length > 0) opts.appProperties = ap;
    }
  }

  // Custom properties
  if (xlsx.customProps) {
    const customPropsEl = xlsx.doc.get(xlsx.customProps);
    if (customPropsEl) {
      const cp = customPropertiesDesc.parse(customPropsEl, {} as ReadContext);
      if (cp.properties?.length) opts.customProperties = cp.properties;
    }
  }

  // Shared strings — use descriptor.parse() for richer data, then extract strings for lookup
  let strings: string[] = [];
  if (xlsx.sharedStrings) {
    const parsed = sharedStringsDesc.parse(xlsx.sharedStrings, {} as never) as {
      entries: (string | Record<string, unknown>)[];
    };
    strings = extractStringsFromEntries(parsed);
  }

  // Create read context for descriptor pipeline
  const readContext = new XlsxReadContext(xlsx, strings);

  // Parse styles (fonts, fills, borders, cellXfs)
  if (xlsx.styles) {
    const parsedStyles = stylesDesc.parse(xlsx.styles, readContext) as unknown as Record<
      string,
      unknown
    >;
    readContext.parsedStyles = parsedStyles;

    // Expose styles sections onto the returned opts for round-trip. compiler.ts
    // re-emits dxfs from options; colors/tableStyles/cellStyles/styleExtensions
    // are surfaced here even though the compiler currently only consumes dxfs,
    // so callers retain the parsed data and the fields stay documented.
    if (parsedStyles.dxfs) opts.dxfs = parsedStyles.dxfs;
    if (parsedStyles.colors) opts.colors = parsedStyles.colors;
    if (parsedStyles.customCellStyles) opts.cellStyles = parsedStyles.customCellStyles;
    if (parsedStyles.styleExtensions) opts.styleExtensions = parsedStyles.styleExtensions;
    const tableStylesInfo = parsedStyles.tableStylesInfo as { tableStyles?: unknown[] } | undefined;
    if (tableStylesInfo?.tableStyles) opts.tableStyles = tableStylesInfo.tableStyles;
  }

  // Parse workbook via descriptor for richer data
  let sheetNames: string[] = [];
  if (xlsx.workbook) {
    const wbData = workbookDesc.parse(xlsx.workbook, readContext) as unknown as Record<
      string,
      unknown
    >;
    const sheets = wbData.sheets as
      | { name: string; sheetId: number; rId: string; state?: string }[]
      | undefined;
    if (sheets) sheetNames = sheets.map((s) => s.name);

    // Workbook-level properties
    if (wbData.protection) opts.workbookProtection = wbData.protection;
    if (wbData.bookView) opts.bookView = wbData.bookView;
    if (wbData.calcPr) opts.calcPr = wbData.calcPr;
    if (wbData.customViews) opts.customWorkbookViews = wbData.customViews;
    if (wbData.fileRecoveryPr) opts.fileRecoveryPr = wbData.fileRecoveryPr;
    if (wbData.functionGroups) opts.functionGroups = wbData.functionGroups;
    if (wbData.webPublishing) opts.webPublishing = wbData.webPublishing;
    if (wbData.fileSharing) opts.fileSharing = wbData.fileSharing;
    if (wbData.workbookPr) opts.workbookPr = wbData.workbookPr;
    if (wbData.volTypes) opts.volTypes = wbData.volTypes;
    if (wbData.webPublishObjects) opts.webPublishObjects = wbData.webPublishObjects;
  }

  // Parse worksheets using descriptor pipeline
  const worksheets: WorksheetOptions[] = [];
  for (let i = 0; i < xlsx.worksheets.length; i++) {
    const wsPath = xlsx.worksheets[i];
    const wsEl = xlsx.doc.get(wsPath);
    if (!wsEl) continue;

    const wsOpts = worksheetDesc.parse(wsEl, readContext) as Record<string, unknown>;
    if (sheetNames[i]) wsOpts.name = sheetNames[i];

    // ── Resolve sub-parts via worksheet relationships ──

    // Comments
    const commentRels = readContext.getWorksheetRelsByType(wsPath, "/comments");
    for (const cr of commentRels) {
      const commentEl = xlsx.doc.get(cr.target);
      if (!commentEl) continue;
      const commentData = commentsDesc.parse(commentEl, readContext) as unknown as Record<
        string,
        unknown
      >;
      if (commentData.comments) {
        wsOpts.comments = commentData.comments;
        break; // one comments file per worksheet
      }
    }

    // Drawings (images + charts)
    const drawingRels = readContext.getWorksheetRelsByType(wsPath, "/drawing");
    for (const dr of drawingRels) {
      const drawingEl = xlsx.doc.get(dr.target);
      if (!drawingEl) continue;
      const drawingData = drawingDesc.parse(drawingEl, readContext);
      if (drawingData.images) wsOpts.images = drawingData.images;
      if (drawingData.charts) wsOpts.charts = drawingData.charts;
      break;
    }

    // Tables
    const tableRels = readContext.getWorksheetRelsByType(wsPath, "/table");
    if (tableRels.length > 0) {
      const tables: Record<string, unknown>[] = [];
      for (const tr of tableRels) {
        const tableEl = xlsx.doc.get(tr.target);
        if (!tableEl) continue;
        const tableData = tableDesc.parse(tableEl, readContext) as unknown as Record<
          string,
          unknown
        >;
        tables.push(tableData);
      }
      if (tables.length > 0) wsOpts.tables = tables;
    }

    // Pivot tables
    const pivotRels = readContext.getWorksheetRelsByType(wsPath, "/pivotTable");
    if (pivotRels.length > 0) {
      const pivotTables: Record<string, unknown>[] = [];
      for (const pr of pivotRels) {
        const pivotEl = xlsx.doc.get(pr.target);
        if (!pivotEl) continue;
        const pivotData = pivotTableDesc.parse(pivotEl, readContext) as unknown as Record<
          string,
          unknown
        >;
        pivotTables.push(pivotData);
      }
      if (pivotTables.length > 0) wsOpts.pivotTables = pivotTables;
    }

    // Resolve external hyperlink URLs
    const hyperlinks = wsOpts.hyperlinks as Record<string, unknown>[] | undefined;
    if (hyperlinks) {
      for (const hl of hyperlinks) {
        const target = hl.target as Record<string, unknown> | undefined;
        if (target?.type === "external" && typeof target.url === "string") {
          const resolved = readContext.resolveWorksheetRel(wsPath, target.url);
          if (resolved) target.url = resolved;
        }
      }
    }

    worksheets.push(wsOpts as WorksheetOptions);
  }

  opts.worksheets = worksheets;

  // Chartsheets — parse chartsheet parts
  const chartsheetPaths = xlsx.doc.keys("xl/chartsheets/").filter((k) => k.endsWith(".xml"));
  if (chartsheetPaths.length > 0) {
    const chartsheets: ChartsheetOptions[] = [];
    for (let i = 0; i < chartsheetPaths.length; i++) {
      const csEl = xlsx.doc.get(chartsheetPaths[i]);
      if (!csEl) continue;
      const csData = chartsheetDesc.parse(csEl, readContext) as unknown as Record<string, unknown>;
      if (sheetNames[worksheets.length + i]) csData.name = sheetNames[worksheets.length + i];
      chartsheets.push(csData as unknown as ChartsheetOptions);
    }
    if (chartsheets.length > 0) opts.chartsheets = chartsheets;
  }

  // Pivot cache definitions and records
  const pivotCacheDefPaths = xlsx.doc
    .keys("xl/pivotCache/")
    .filter((k) => k.includes("pivotCacheDefinition"));
  if (pivotCacheDefPaths.length > 0) {
    const pivotCaches: Record<string, unknown>[] = [];
    for (const pcdPath of pivotCacheDefPaths) {
      const pcdEl = xlsx.doc.get(pcdPath);
      if (!pcdEl) continue;
      const pcdData = pivotCacheDefDesc.parse(pcdEl, readContext) as unknown as Record<
        string,
        unknown
      >;
      pivotCaches.push(pcdData);
    }
    if (pivotCaches.length > 0) opts.pivotCaches = pivotCaches;
  }

  const pivotCacheRecPaths = xlsx.doc
    .keys("xl/pivotCache/")
    .filter((k) => k.includes("pivotCacheRecords"));
  if (pivotCacheRecPaths.length > 0) {
    const pivotCacheRecords: Record<string, unknown>[] = [];
    for (const pcrPath of pivotCacheRecPaths) {
      const pcrEl = xlsx.doc.get(pcrPath);
      if (!pcrEl) continue;
      const pcrData = pivotCacheRecordsDesc.parse(pcrEl, readContext) as unknown as Record<
        string,
        unknown
      >;
      pivotCacheRecords.push(pcrData);
    }
    if (pivotCacheRecords.length > 0) opts.pivotCacheRecords = pivotCacheRecords;
  }

  // Calculation chain
  const calcChainEl = xlsx.doc.get("xl/calcChain.xml");
  if (calcChainEl) {
    const calcData = calcChainDesc.parse(calcChainEl, readContext) as unknown as Record<
      string,
      unknown
    >;
    if (calcData.cells) opts.calcChain = calcData.cells;
  }

  // External links
  const extLinkPaths = xlsx.doc.keys("xl/externalLinks/").filter((k) => k.endsWith(".xml"));
  if (extLinkPaths.length > 0) {
    const externalLinks: Record<string, unknown>[] = [];
    for (const elPath of extLinkPaths) {
      const elEl = xlsx.doc.get(elPath);
      if (!elEl) continue;
      const elData = externalLinkDesc.parse(elEl, readContext) as Record<string, unknown>;

      // Resolve the external book target from the sibling rels file
      // (xl/externalLinks/_rels/externalLinkN.xml.rels), which compiler.ts writes.
      const elData2 = elData as { externalBook?: { target?: string } };
      if (elData2.externalBook) {
        const relsPath = externalLinkRelsPath(elPath);
        const relsEl = xlsx.doc.get(relsPath);
        if (relsEl) {
          for (const child of relsEl.elements ?? []) {
            if (child.name !== "Relationship") continue;
            const type = attr(child, "Type") ?? "";
            if (!type.includes("/externalLinkPath")) continue;
            const target = attr(child, "Target");
            if (target) {
              elData2.externalBook.target = target;
              break;
            }
          }
        }
      }

      externalLinks.push(elData);
    }
    if (externalLinks.length > 0) opts.externalLinks = externalLinks;
  }

  // Shared-workbook revisions: workbook.xml.rels → revisionHeaders/users;
  // revisionHeaders.xml.rels → per-header revision logs.
  const wbRelsEl2 = xlsx.doc.get("xl/_rels/workbook.xml.rels");
  let revHeadersTarget: string | undefined;
  let usersTarget: string | undefined;
  if (wbRelsEl2) {
    for (const child of wbRelsEl2.elements ?? []) {
      if (child.name !== "Relationship") continue;
      const type = attr(child, "Type") ?? "";
      const target = attr(child, "Target") ?? "";
      if (type.includes("/revisionHeaders")) revHeadersTarget = target;
      else if (type.includes("/users")) usersTarget = target;
    }
  }
  if (revHeadersTarget) {
    const headersPath = revHeadersTarget.startsWith("/")
      ? revHeadersTarget.slice(1)
      : `xl/${revHeadersTarget}`;
    const headersEl = xlsx.doc.get(headersPath);
    if (headersEl) {
      const headers = revisionHeadersDesc.parse(headersEl, readContext);
      const logs: RevisionLogOptions[] = [];
      const slash = headersPath.lastIndexOf("/");
      const revHeadersRelsEl = xlsx.doc.get(
        `${headersPath.slice(0, slash)}/_rels/${headersPath.slice(slash + 1)}.rels`,
      );
      if (revHeadersRelsEl) {
        for (const child of revHeadersRelsEl.elements ?? []) {
          if (child.name !== "Relationship") continue;
          if (!(attr(child, "Type") ?? "").includes("/revisionLog")) continue;
          const t = attr(child, "Target");
          if (!t) continue;
          const logEl = xlsx.doc.get(t.startsWith("/") ? t.slice(1) : `xl/${t}`);
          if (logEl) logs.push(revisionLogDesc.parse(logEl, readContext));
        }
      }
      const revisionLog: SharedWorkbookOptions = { headers, logs };
      if (usersTarget) {
        const usersEl = xlsx.doc.get(
          usersTarget.startsWith("/") ? usersTarget.slice(1) : `xl/${usersTarget}`,
        );
        if (usersEl) {
          const users = usersDesc.parse(usersEl, readContext);
          if (users.users) revisionLog.users = users;
        }
      }
      opts.revisionLog = revisionLog;
    }
  }

  return opts as WorkbookOptions;
}
