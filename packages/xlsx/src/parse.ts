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
import { partPathToRelsPath, toUint8Array } from "@office-open/core";
import type { DataType } from "@office-open/core";
import type { ReadContext } from "@office-open/core/descriptor";
import { themeDesc } from "@office-open/core/theme";
import type { Element } from "@office-open/xml";
import type { ParseOptions } from "@office-open/xml";
import { attr } from "@office-open/xml";
import { calcChainDesc } from "@parts/calc-chain";
import { chartsheetDesc } from "@parts/chartsheet";
import type { ChartsheetOptions } from "@parts/chartsheet";
import { commentsDesc } from "@parts/comments";
import { connectionsDesc } from "@parts/connection";
import { dialogsheetDesc } from "@parts/dialogsheet";
import type { DialogsheetOptions } from "@parts/dialogsheet";
import { drawingDesc } from "@parts/drawing";
import { externalLinkDesc } from "@parts/external-link";
import type { ExternalLinkOptions } from "@parts/external-link";
import type { SharedWorkbookOptions, WorkbookOptions } from "@parts/file";
import { metadataDesc } from "@parts/metadata";
import { pivotCacheDefDesc, pivotCacheRecordsDesc } from "@parts/pivot-cache";
import type { PivotCacheDefParseResult, PivotCacheRecordsParseResult } from "@parts/pivot-cache";
import { pivotTableDesc } from "@parts/pivot-table";
import type { PivotTableParseResult } from "@parts/pivot-table";
import type { PivotTableOptions } from "@parts/pivot/pivot-utils";
import { queryTableDesc } from "@parts/query-table";
import type { QueryTableOptions } from "@parts/query-table";
import {
  revisionHeadersDesc,
  revisionLogDesc,
  usersDesc,
  type RevisionLogOptions,
} from "@parts/revision-log";
import { sharedStringsDesc } from "@parts/shared-strings";
import type { SharedStringsDocOptions } from "@parts/shared-strings";
import { stylesDesc } from "@parts/styles";
import { tableDesc } from "@parts/table";
import type { TableOptions } from "@parts/table";
import { workbookDesc } from "@parts/workbook";
import { worksheetDesc } from "@parts/worksheet";
import type { WorksheetChartOptions, PictureOptions, WorksheetOptions } from "@parts/worksheet";
import { mapInfoDesc, singleXmlCellsDesc } from "@parts/xml-mapping";
import type { SingleXmlCellOptions } from "@parts/xml-mapping";

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
  /** xl/theme/theme{n}.xml path (resolved from workbook rels) */
  theme?: string;
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

/**
 * Worksheet parts read with sheetData deferred — the XML parser captures the
 * container's inner XML verbatim and `parseSheetDataRows` walks it directly.
 */
const WORKSHEET_PARSE_OPTIONS: ParseOptions = { deferElements: ["sheetData"] };

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
  let theme: string | undefined;

  const wbRels = doc.get("xl/_rels/workbook.xml.rels");
  if (wbRels) {
    for (const child of wbRels.elements ?? []) {
      if (child.name !== "Relationship") continue;
      const type = attr(child, "Type") ?? "";
      const target = attr(child, "Target") ?? "";
      if (!target) continue;

      if (type.includes("/worksheet")) {
        worksheets.push(target.startsWith("/") ? target.slice(1) : `xl/${target}`);
      } else if (type.includes("/theme")) {
        theme = target.startsWith("/") ? target.slice(1) : `xl/${target}`;
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
    theme,
    partRefs: { worksheets, charts, media, drawings },
    coreProps,
    appProps,
    customProps,
  };
}

// ── Shared strings helper ──

/** Extract plain string array from sharedStringsDesc.parse() result. */
function extractStringsFromEntries(parsed: SharedStringsDocOptions): string[] {
  const strings: string[] = [];
  for (const entry of parsed.entries) {
    if (typeof entry === "string") {
      strings.push(entry);
    } else if (entry.runs && entry.runs.length > 0) {
      // Rich text: concatenate run texts
      strings.push(entry.runs.map((r) => r.text).join(""));
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

  const opts: Partial<WorkbookOptions> = {};

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
    const parsed = sharedStringsDesc.parse(xlsx.sharedStrings, {} as never);
    strings = extractStringsFromEntries(parsed);
  }

  // Create read context for descriptor pipeline
  const readContext = new XlsxReadContext(xlsx, strings);

  // Parse styles (fonts, fills, borders, cellXfs)
  if (xlsx.styles) {
    const parsedStyles = stylesDesc.parse(xlsx.styles, readContext);
    readContext.parsedStyles = parsedStyles;

    // Expose styles sections onto the returned opts for round-trip. compiler.ts
    // re-emits dxfs from options; colors/tableStyles/cellStyles/styleExtensions
    // are surfaced here even though the compiler currently only consumes dxfs,
    // so callers retain the parsed data and the fields stay documented.
    if (parsedStyles.dxfs) opts.dxfs = parsedStyles.dxfs;
    if (parsedStyles.colors) opts.colors = parsedStyles.colors;
    if (parsedStyles.customCellStyles) opts.cellStyles = parsedStyles.customCellStyles;
    if (parsedStyles.cellStyleXfs) opts.cellStyleXfs = parsedStyles.cellStyleXfs;
    if (parsedStyles.styleExtensions) opts.styleExtensions = parsedStyles.styleExtensions;
    if (parsedStyles.tableStylesInfo?.tableStyles)
      opts.tableStyles = parsedStyles.tableStylesInfo.tableStyles;
  }

  // Theme — structured round-trip so a custom source theme survives instead of
  // being replaced by the compiler's fresh default.
  if (xlsx.theme) {
    const themeEl = xlsx.doc.get(xlsx.theme);
    if (themeEl) {
      const themeOptions = themeDesc.parse(themeEl, readContext);
      if (themeOptions) opts.theme = themeOptions;
    }
  }

  // Parse workbook via descriptor for richer data
  let sheetNames: string[] = [];
  if (xlsx.workbook) {
    const wbData = workbookDesc.parse(xlsx.workbook, readContext);
    if (wbData.sheets) sheetNames = wbData.sheets.map((s) => s.name);

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
    if (wbData.definedNames) opts.definedNames = wbData.definedNames;
  }

  // Parse worksheets using descriptor pipeline. Worksheet parts defer
  // sheetData — the row scanner walks the captured inner XML, skipping the
  // per-cell Element tree (the dominant allocation cost on large sheets).
  const worksheets: WorksheetOptions[] = [];
  for (const [i, wsPath] of xlsx.worksheets.entries()) {
    const wsEl = xlsx.doc.get(wsPath, WORKSHEET_PARSE_OPTIONS);
    if (!wsEl) continue;

    const wsOpts = worksheetDesc.parse(wsEl, readContext);
    if (sheetNames[i]) wsOpts.name = sheetNames[i];

    // ── Resolve sub-parts via worksheet relationships ──

    // Comments
    const commentRels = readContext.getWorksheetRelsByType(wsPath, "/comments");
    for (const cr of commentRels) {
      const commentEl = xlsx.doc.get(cr.target);
      if (!commentEl) continue;
      const commentData = commentsDesc.parse(commentEl, readContext);
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
      // drawingDesc.parse yields CT-layer shapes (rId-anchored DrawingImage/
      // DrawingChart), while WorksheetOptions.images/charts are the user-layer
      // shapes (data bytes / ChartSpace content). The shapes are not
      // interchangeable — round-trip drawings are lossy — these casts mark the
      // known impedance, same as wsOpts.pivotTables above.
      if (drawingData.images) wsOpts.images = drawingData.images as unknown as PictureOptions[];
      if (drawingData.charts)
        wsOpts.charts = drawingData.charts as unknown as WorksheetChartOptions[];
      // Shapes/connectors/groups pass through unchanged (no media bridge).
      if (drawingData.shapes) wsOpts.shapes = drawingData.shapes;
      if (drawingData.connectors) wsOpts.connectors = drawingData.connectors;
      if (drawingData.groups) wsOpts.groups = drawingData.groups;
      break;
    }

    // Tables
    const tableRels = readContext.getWorksheetRelsByType(wsPath, "/table");
    if (tableRels.length > 0) {
      const tables: TableOptions[] = [];
      for (const tr of tableRels) {
        const tableEl = xlsx.doc.get(tr.target);
        if (!tableEl) continue;
        const tableData = tableDesc.parse(tableEl, readContext);
        tables.push(tableData);
      }
      if (tables.length > 0) wsOpts.tables = tables;
    }

    // Query tables
    const queryTableRels = readContext.getWorksheetRelsByType(wsPath, "/queryTable");
    if (queryTableRels.length > 0) {
      const queryTables: QueryTableOptions[] = [];
      for (const qtr of queryTableRels) {
        const qtEl = xlsx.doc.get(qtr.target);
        if (!qtEl) continue;
        queryTables.push(queryTableDesc.parse(qtEl, readContext));
      }
      if (queryTables.length > 0) wsOpts.queryTables = queryTables;
    }

    // Single-cell XML tables
    const singleXmlCellRels = readContext.getWorksheetRelsByType(wsPath, "/tableSingleCells");
    if (singleXmlCellRels.length > 0) {
      const singleXmlCells: SingleXmlCellOptions[] = [];
      for (const sxr of singleXmlCellRels) {
        const sxEl = xlsx.doc.get(sxr.target);
        if (!sxEl) continue;
        singleXmlCells.push(...singleXmlCellsDesc.parse(sxEl, readContext).cells);
      }
      if (singleXmlCells.length > 0) wsOpts.singleXmlCells = singleXmlCells;
    }

    // Pivot tables
    const pivotRels = readContext.getWorksheetRelsByType(wsPath, "/pivotTable");
    if (pivotRels.length > 0) {
      const pivotTables: PivotTableParseResult[] = [];
      for (const pr of pivotRels) {
        const pivotEl = xlsx.doc.get(pr.target);
        if (!pivotEl) continue;
        const pivotData = pivotTableDesc.parse(pivotEl, readContext);
        pivotTables.push(pivotData);
      }
      // WorksheetOptions.pivotTables is PivotTableOptions[] (user-layer) for the
      // compiler's stringify; parse yields the CT-layer PivotTableParseResult.
      // The shapes are not interchangeable (index-based vs name-based), so
      // round-trip pivot tables are lossy — this cast marks the known impedance.
      if (pivotTables.length > 0)
        wsOpts.pivotTables = pivotTables as unknown as PivotTableOptions[];
    }

    // Resolve external hyperlink URLs
    const hyperlinks = wsOpts.hyperlinks;
    if (hyperlinks) {
      for (const hl of hyperlinks) {
        if (hl.target.type === "external") {
          const resolved = readContext.resolveWorksheetRel(wsPath, hl.target.url);
          if (resolved) hl.target.url = resolved;
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
    for (const [i, csPath] of chartsheetPaths.entries()) {
      const csEl = xlsx.doc.get(csPath);
      if (!csEl) continue;
      const csData = chartsheetDesc.parse(csEl, readContext);
      if (sheetNames[worksheets.length + i]) csData.name = sheetNames[worksheets.length + i];
      chartsheets.push(csData);
    }
    if (chartsheets.length > 0) opts.chartsheets = chartsheets;
  }

  // Dialogsheets — parse legacy dialog sheet parts
  const dialogsheetPaths = xlsx.doc.keys("xl/dialogSheets/").filter((k) => k.endsWith(".xml"));
  if (dialogsheetPaths.length > 0) {
    const dialogsheets: DialogsheetOptions[] = [];
    for (const dsPath of dialogsheetPaths) {
      const dsEl = xlsx.doc.get(dsPath);
      if (!dsEl) continue;
      const dsData = dialogsheetDesc.parse(dsEl, readContext);
      dialogsheets.push(dsData);
    }
    if (dialogsheets.length > 0) opts.dialogsheets = dialogsheets;
  }

  // Pivot cache definitions and records
  const pivotCacheDefPaths = xlsx.doc
    .keys("xl/pivotCache/")
    .filter((k) => k.includes("pivotCacheDefinition"));
  if (pivotCacheDefPaths.length > 0) {
    const pivotCaches: PivotCacheDefParseResult[] = [];
    for (const pcdPath of pivotCacheDefPaths) {
      const pcdEl = xlsx.doc.get(pcdPath);
      if (!pcdEl) continue;
      const pcdData = pivotCacheDefDesc.parse(pcdEl, readContext);
      pivotCaches.push(pcdData);
    }
    if (pivotCaches.length > 0) opts.pivotCaches = pivotCaches;
  }

  const pivotCacheRecPaths = xlsx.doc
    .keys("xl/pivotCache/")
    .filter((k) => k.includes("pivotCacheRecords"));
  if (pivotCacheRecPaths.length > 0) {
    const pivotCacheRecords: PivotCacheRecordsParseResult[] = [];
    for (const pcrPath of pivotCacheRecPaths) {
      const pcrEl = xlsx.doc.get(pcrPath);
      if (!pcrEl) continue;
      const pcrData = pivotCacheRecordsDesc.parse(pcrEl, readContext);
      pivotCacheRecords.push(pcrData);
    }
    if (pivotCacheRecords.length > 0) opts.pivotCacheRecords = pivotCacheRecords;
  }

  // Calculation chain
  const calcChainEl = xlsx.doc.get("xl/calcChain.xml");
  if (calcChainEl) {
    const calcData = calcChainDesc.parse(calcChainEl, readContext);
    if (calcData.cells) opts.calcChain = calcData.cells;
  }

  // Connections (xl/connections.xml)
  const connectionsEl = xlsx.doc.get("xl/connections.xml");
  if (connectionsEl) {
    const connData = connectionsDesc.parse(connectionsEl, readContext);
    if (connData.connections.length > 0) opts.connections = connData.connections;
  }

  // Rich metadata (xl/metadata.xml)
  const metadataEl = xlsx.doc.get("xl/metadata.xml");
  if (metadataEl) {
    const metadataData = metadataDesc.parse(metadataEl, readContext);
    opts.metadata = metadataData;
  }

  // XML mappings (xl/xmlMaps.xml)
  const xmlMapsEl = xlsx.doc.get("xl/xmlMaps.xml");
  if (xmlMapsEl) {
    opts.xmlMaps = mapInfoDesc.parse(xmlMapsEl, readContext);
  }

  // External links
  const extLinkPaths = xlsx.doc.keys("xl/externalLinks/").filter((k) => k.endsWith(".xml"));
  if (extLinkPaths.length > 0) {
    const externalLinks: ExternalLinkOptions[] = [];
    for (const elPath of extLinkPaths) {
      const elEl = xlsx.doc.get(elPath);
      if (!elEl) continue;
      const elData = externalLinkDesc.parse(elEl, readContext);

      // Resolve the external book target from the sibling rels file
      // (xl/externalLinks/_rels/externalLinkN.xml.rels), which compiler.ts writes.
      if (elData.externalBook) {
        const relsPath = partPathToRelsPath(elPath);
        const relsEl = xlsx.doc.get(relsPath);
        if (relsEl) {
          for (const child of relsEl.elements ?? []) {
            if (child.name !== "Relationship") continue;
            const type = attr(child, "Type") ?? "";
            if (!type.includes("/externalLinkPath")) continue;
            const target = attr(child, "Target");
            if (target) {
              elData.externalBook.target = target;
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
      const revHeadersRelsEl = xlsx.doc.get(partPathToRelsPath(headersPath));
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
