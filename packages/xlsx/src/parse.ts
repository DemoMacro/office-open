/**
 * XLSX parsing — parse .xlsx files into structured data.
 *
 * @module
 */
import {
  appPropertiesDesc,
  contentTypesDesc,
  convertToEmu,
  customPropertiesDesc,
  parseArchive,
  parseCorePropsElement,
  ParsedArchive,
} from "@office-open/core";
import {
  collectPassthroughParts,
  isEncryptedContainer,
  opaquePassthroughPolicy,
  partPathToRelsPath,
  pickNonVisualDrawingProperties,
  resolveRelationshipTarget,
  toUint8Array,
  toUint8ArrayAsync,
} from "@office-open/core";
import type { DataType } from "@office-open/core";
import { chartSpaceDesc, userShapesDesc, type ExternalDataOptions } from "@office-open/core/chart";
import type { ReadContext } from "@office-open/core/descriptor";
import { themeDesc } from "@office-open/core/theme";
import type { Element } from "@office-open/xml";
import type { ParseOptions } from "@office-open/xml";
import { attr, attrNum, findChild } from "@office-open/xml";
import { ActiveXControlParseError, activeXControlDesc } from "@parts/active-x-control";
import { calcChainDesc } from "@parts/calc-chain";
import { chartsheetDesc } from "@parts/chartsheet";
import type { ChartsheetOptions } from "@parts/chartsheet";
import { commentsDesc, mergeNoteAnchors, vmlNotesDesc } from "@parts/comments";
import { connectionsDesc } from "@parts/connection";
import { ControlPropertiesParseError, controlPropertiesDesc } from "@parts/control-properties";
import { dialogsheetDesc } from "@parts/dialogsheet";
import type { DialogsheetOptions } from "@parts/dialogsheet";
import { drawingDesc, pickAnchorOptions } from "@parts/drawing";
import { externalLinkDesc } from "@parts/external-link";
import type { ExternalLinkOptions } from "@parts/external-link";
import type { SharedWorkbookOptions, WorkbookOptions } from "@parts/file";
import type { DefinitionPivotCacheOptions } from "@parts/file";
import { metadataDesc } from "@parts/metadata";
import { parsePivotCacheDefinition } from "@parts/pivot-cache-definition";
import { parsePivotCacheRecords } from "@parts/pivot-cache-records";
import { parsePivotTableDefinition } from "@parts/pivot-table";
import { queryTableDesc } from "@parts/query-table";
import type { QueryTableOptions } from "@parts/query-table";
import {
  revisionHeadersDesc,
  revisionLogDesc,
  usersDesc,
  type RevisionLogOptions,
} from "@parts/revision-log";
import { sharedStringsDesc } from "@parts/shared-strings";
import { stylesDesc } from "@parts/styles";
import { tableDesc } from "@parts/table";
import type { TableOptions } from "@parts/table";
import { parseVolTypesEl } from "@parts/vol-types";
import { workbookDesc } from "@parts/workbook";
import type { RichTextOptions } from "@parts/worksheet";
import { worksheetDesc } from "@parts/worksheet";
import type {
  WorksheetChartOptions,
  WorksheetWebExtensionOptions,
  WorksheetSmartArtOptions,
  PictureOptions,
  WorksheetContentPartOptions,
  WorksheetOptions,
} from "@parts/worksheet";
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
  /** Resolved primary workbook part path (xl/workbook.xml by default) */
  workbookPath: string;
  /** Primary workbook root element */
  workbook?: Element;
  /** Worksheet paths (xl/worksheets/sheet{n}.xml) */
  worksheets: string[];
  /** xl/styles.xml root element */
  styles?: Element;
  /** Resolved styles part path (xl/styles.xml by default) */
  stylesPath: string;
  /** xl/sharedStrings.xml root element */
  sharedStrings?: Element;
  /** Resolved sharedStrings part path (xl/sharedStrings.xml by default) */
  sharedStringsPath: string;
  /** xl/theme/theme{n}.xml path (resolved from workbook rels) */
  theme?: string;
  partRefs: XlsxPartRefs;
  /** docProps/core.xml path */
  coreProps?: string;
  /** Root-rels core-properties relationship type (source lexical form). */
  corePropertiesType?: string;
  /** docProps/app.xml path */
  appProps?: string;
  /** docProps/custom.xml path */
  customProps?: string;
  /** Legacy Microsoft OPC relationship namespace flavor detected in manifests. */
  packageRelationshipNamespace?: "microsoft2005";
}

const LEADING_PATH_NUMBER = /(\d+)/;

function sortByNumber(paths: string[]): string[] {
  // Decorate-sort: extract each path's number once instead of per comparison.
  return paths
    .map((p) => ({ p, n: parseInt(p.match(LEADING_PATH_NUMBER)?.[1] ?? "0", 10) }))
    .sort((a, b) => a.n - b.n)
    .map(({ p }) => p);
}

function graphicDataUris(element: Element, uris: string[] = []): string[] {
  if (element.name === "a:graphicData") {
    const uri = attr(element, "uri");
    if (uri !== undefined) uris.push(uri);
  }
  for (const child of element.elements ?? []) graphicDataUris(child, uris);
  return uris;
}

function isLegacyChartDrawing(element: Element | undefined): boolean {
  if (!element) return false;
  const uris = graphicDataUris(element);
  return (
    uris.some((uri) =>
      uri.startsWith("http://schemas.microsoft.com/office/excel/2005/8/ChartML"),
    ) && !uris.some((uri) => uri.includes("openxmlformats.org/drawingml/2006/chart"))
  );
}

/**
 * Fill a parsed chart's userShapes anchors from the companion part body —
 * chartSpaceDesc reads only the c:userShapes r:id; the body hangs off the
 * chart part's own rels (chartUserShapes relationship).
 */
function readChartUserShapes(
  chartPath: string | undefined,
  chart: {
    userShapes?: {
      relationshipId?: string;
      anchors: unknown[];
      path?: string;
      rootElement?: "chart" | "drawing";
    };
  },
  readContext: XlsxReadContext,
  doc: XlsxDocument["doc"],
): void {
  const rid = chart.userShapes?.relationshipId;
  if (rid === undefined || chartPath === undefined) return;
  const rel = readContext
    .getWorksheetRelsByType(chartPath, "/chartUserShapes")
    .find((r) => r.rId === rid);
  const bodyEl = rel ? doc.get(rel.target) : undefined;
  if (!bodyEl) return;
  const body = userShapesDesc.parse(bodyEl, readContext);
  chart.userShapes = {
    ...chart.userShapes,
    rootElement: body.rootElement,
    anchors: body.anchors,
    ...(rel ? { path: rel.target } : {}),
  };
}

function readChartExternalLink(
  chartPath: string,
  externalData: ExternalDataOptions | undefined,
  readContext: XlsxReadContext,
  doc: XlsxDocument["doc"],
  chartExternalLinkPaths: Set<string>,
): {
  externalLink?: ExternalLinkOptions;
  externalLinkPath?: string;
  externalDataRelationshipType?: string;
  externalDataRelationshipTarget?: string;
} {
  const chartRelsEl = doc.get(partPathToRelsPath(chartPath));
  const externalDataRel = chartRelsEl?.elements?.find(
    (rel) => rel.name === "Relationship" && attr(rel, "Id") === externalData?.relationshipId,
  );
  const externalDataTarget = externalDataRel ? attr(externalDataRel, "Target") : undefined;
  if (!externalDataRel || !externalDataTarget) return {};

  const externalDataRelationshipType = attr(externalDataRel, "Type");
  const result: {
    externalLink?: ExternalLinkOptions;
    externalLinkPath?: string;
    externalDataRelationshipType?: string;
    externalDataRelationshipTarget?: string;
  } = {
    externalDataRelationshipType,
    externalDataRelationshipTarget: externalDataTarget,
  };
  if (!externalDataRelationshipType?.includes("/externalLinkPath")) return result;

  const externalLinkPath = resolveRelationshipTarget(chartPath, externalDataTarget);
  const externalLinkEl = doc.get(externalLinkPath);
  if (!externalLinkEl) return result;

  const externalLink = externalLinkDesc.parse(externalLinkEl, readContext);
  const linkRelsEl = doc.get(partPathToRelsPath(externalLinkPath));
  const bookRel = linkRelsEl?.elements?.find(
    (rel) => rel.name === "Relationship" && (attr(rel, "Type") ?? "").includes("/externalLinkPath"),
  );
  const bookTarget = bookRel ? attr(bookRel, "Target") : undefined;
  if (bookTarget && externalLink.externalBook) externalLink.externalBook.target = bookTarget;
  chartExternalLinkPaths.add(externalLinkPath);
  return {
    ...result,
    externalLink,
    externalLinkPath,
  };
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
  return parseXlsxArchive(parseArchive(toUint8Array(data)));
}

/**
 * Resolve the primary workbook part from `_rels/.rels` (OPC officeDocument
 * relationship). Minimal hand-built packages may keep every part at the root.
 */
function resolveWorkbookPath(doc: ParsedArchive): string {
  const rootRels = doc.get("_rels/.rels");
  for (const child of rootRels?.elements ?? []) {
    if (child.name !== "Relationship") continue;
    if (!(attr(child, "Type") ?? "").endsWith("/officeDocument")) continue;
    if (attr(child, "TargetMode") === "External") continue;
    const target = attr(child, "Target") ?? "";
    if (!target) continue;
    const resolved = target.startsWith("/") ? target.slice(1) : target.replaceAll("\\", "/");
    if (doc.get(resolved)) return resolved;
  }
  return "xl/workbook.xml";
}

/** Directory prefix of a part path ("" for root-level parts). */
function dirOf(partPath: string): string {
  const slash = partPath.lastIndexOf("/");
  return slash === -1 ? "" : partPath.slice(0, slash);
}

/** Resolve a rel target relative to the workbook's directory. */
function resolveWorkbookTarget(target: string, dir: string): string {
  if (target.startsWith("/")) return target.slice(1);
  return dir ? `${dir}/${target}` : target;
}

/** Archive-backed core of {@link parseXlsx} — shared with the Blob open path. */
function parseXlsxArchive(doc: ParsedArchive): XlsxDocument {
  const workbookPath = resolveWorkbookPath(doc);
  const wbDir = dirOf(workbookPath);
  const workbook = doc.get(workbookPath);

  // Resolve worksheet paths from workbook rels
  let worksheets: string[] = [];
  let charts: string[] = [];
  let drawings: string[] = [];
  const media: string[] = [];
  let theme: string | undefined;

  const wbRels = doc.get(partPathToRelsPath(workbookPath));
  let stylesPath = "xl/styles.xml";
  let sharedStringsPath = "xl/sharedStrings.xml";
  if (wbRels) {
    for (const child of wbRels.elements ?? []) {
      if (child.name !== "Relationship") continue;
      const type = attr(child, "Type") ?? "";
      const target = attr(child, "Target") ?? "";
      if (!target) continue;

      if (type.includes("/worksheet")) {
        worksheets.push(resolveWorkbookTarget(target, wbDir));
      } else if (type.includes("/theme")) {
        theme = resolveWorkbookTarget(target, wbDir);
      } else if (type.endsWith("/styles")) {
        stylesPath = resolveWorkbookTarget(target, wbDir);
      } else if (type.endsWith("/sharedStrings")) {
        sharedStringsPath = resolveWorkbookTarget(target, wbDir);
      }
    }
  }
  const styles = doc.get(stylesPath);
  const sharedStrings = doc.get(sharedStringsPath);
  worksheets = sortByNumber(worksheets);

  // Scan for drawings, charts, media
  drawings.push(...doc.keys("xl/drawings/").filter((k) => k.endsWith(".xml")));
  // Only chartSpace parts drive compilation; style/color companions and
  // userShapes stay in the package passthrough set.
  charts.push(
    ...doc.keys("xl/charts/").filter((k) => {
      if (!k.endsWith(".xml")) return false;
      const chart = doc.get(k);
      const name = chart?.name ?? "";
      const chartNamespace = attr(chart, "xmlns:c") ?? attr(chart, "xmlns");
      return (
        (name === "chartSpace" || name === "c:chartSpace" || name.endsWith(":chartSpace")) &&
        (chartNamespace === undefined ||
          chartNamespace.includes("openxmlformats.org/drawingml/2006/chart") ||
          chartNamespace.includes("purl.oclc.org/ooxml/drawingml/chart"))
      );
    }),
  );
  media.push(...doc.keys("xl/media/"));
  drawings = drawings.filter((path) => !isLegacyChartDrawing(doc.get(path)));
  drawings = sortByNumber(drawings);
  charts = sortByNumber(charts);

  // Root rels → core/app props
  let coreProps: string | undefined;
  let corePropertiesType: string | undefined;
  let appProps: string | undefined;
  let customProps: string | undefined;
  const rootRels = doc.get("_rels/.rels");
  const workbookRels = doc.get(partPathToRelsPath(workbookPath));
  const isLegacyRelationshipNamespace =
    attr(rootRels, "xmlns") === "http://schemas.microsoft.com/package/2005/06/relationships" ||
    attr(workbookRels, "xmlns") === "http://schemas.microsoft.com/package/2005/06/relationships";
  if (rootRels) {
    for (const child of rootRels.elements ?? []) {
      if (child.name !== "Relationship") continue;
      const type = attr(child, "Type") ?? "";
      const target = attr(child, "Target") ?? "";
      if (!target) continue;
      const path = target.startsWith("/") ? target.slice(1) : target;
      // Transitional packages use the oclc URI form with camelCase segments
      // (…/extendedProperties); normalize case and hyphens so both resolve.
      const relType = type.toLowerCase().replaceAll("-", "");
      if (relType.includes("/coreproperties")) coreProps = path;
      if (relType.includes("/coreproperties")) corePropertiesType = type;
      else if (relType.includes("/extendedproperties") || relType.endsWith("/docpropsapp"))
        appProps = path;
      else if (relType.includes("/customproperties")) customProps = path;
    }
  }

  return {
    doc,
    workbookPath,
    workbook,
    worksheets,
    styles,
    stylesPath,
    sharedStrings,
    sharedStringsPath,
    theme,
    partRefs: { worksheets, charts, media, drawings },
    coreProps,
    corePropertiesType,
    appProps,
    customProps,
    ...(isLegacyRelationshipNamespace
      ? { packageRelationshipNamespace: "microsoft2005" as const }
      : {}),
  };
}

// ── Shared strings helper ──

/**
 * Parse a .xlsx file and convert it into WorkbookOptions.
 *
 * Async so that `Blob` (including `File`) and `ReadableStream` inputs are
 * accepted alongside raw bytes — parts are indexed and decompressed on
 * demand. The returned options can be passed to `new Workbook(parsed)`.
 */
export async function parseWorkbook(data: DataType): Promise<WorkbookOptions> {
  // Blob/File inputs bypass full materialization: the archive indexes and
  // inflates through random-access windows, so a multi-GB package costs only
  // the parts actually read.
  if (data instanceof Blob) {
    const head = new Uint8Array(await data.slice(0, 8).arrayBuffer());
    if (isEncryptedContainer(head)) {
      return { encrypted: { data: await toUint8ArrayAsync(data) } };
    }
    return parseWorkbookFromXlsx(parseXlsxArchive(await ParsedArchive.open(data)));
  }
  return parseWorkbookFromBytes(await toUint8ArrayAsync(data));
}

/**
 * Synchronous counterpart of {@link parseWorkbook}: same result, but `Blob`
 * and `ReadableStream` inputs throw (normalize them to bytes first, or use
 * the async entry).
 */
export function parseWorkbookSync(data: DataType): WorkbookOptions {
  return parseWorkbookFromBytes(toUint8Array(data));
}

function parseWorkbookFromBytes(uint8: Uint8Array): WorkbookOptions {
  // Encrypted package (OLE2/CFB container): the plaintext needs the password,
  // so carry the source bytes verbatim for generate() to re-emit.
  if (isEncryptedContainer(uint8)) {
    return { encrypted: { data: uint8 } };
  }

  return parseWorkbookFromXlsx(parseXlsx(uint8));
}

function parseWorkbookFromXlsx(xlsx: XlsxDocument): WorkbookOptions {
  const opts: Partial<WorkbookOptions> = {};

  // Core properties
  if (xlsx.coreProps) {
    const corePropsEl = xlsx.doc.get(xlsx.coreProps);
    if (corePropsEl) {
      const cp = parseCorePropsElement(corePropsEl);
      // Empty strings are meaningful (element present, text empty) — assign
      // the whole shape so they survive round-trip.
      Object.assign(opts, cp);
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

  // Shared strings — rich-text entries flow through as objects so both the
  // cell lookup and the rebuilt table keep their structure; generate() seeds
  // the write context from opts.sharedStrings to preserve si indices.
  let sstEntries: (string | RichTextOptions)[] = [];
  let sharedStringsCount: number | undefined;
  let sharedStringsUniqueCount: number | undefined;
  if (xlsx.sharedStrings) {
    sstEntries = sharedStringsDesc.parse(xlsx.sharedStrings, {} as never).entries;
    sharedStringsCount =
      attrNum(xlsx.sharedStrings, "count") ?? attrNum(xlsx.sharedStrings, "totalCount");
    sharedStringsUniqueCount = attrNum(xlsx.sharedStrings, "uniqueCount");
  }
  if (sstEntries.length > 0) opts.sharedStrings = sstEntries;
  if (xlsx.sharedStrings) opts.sharedStringsDeclared = true;
  if (sharedStringsCount !== undefined) opts.sharedStringsCount = sharedStringsCount;
  if (sharedStringsUniqueCount !== undefined)
    opts.sharedStringsUniqueCount = sharedStringsUniqueCount;
  if (xlsx.coreProps) opts.corePropertiesPath = xlsx.coreProps;
  if (
    xlsx.corePropertiesType !== undefined &&
    xlsx.corePropertiesType !==
      "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
  ) {
    opts.corePropertiesRelationshipType = xlsx.corePropertiesType;
  }
  if (xlsx.appProps) opts.appPropertiesPath = xlsx.appProps;
  if (xlsx.customProps) {
    opts.customPropertiesPath = xlsx.customProps;
    opts.customPropertiesDeclared = true;
  }

  // Create read context for descriptor pipeline
  const readContext = new XlsxReadContext(xlsx, sstEntries);

  // Pivot caches are canonical model entries; records are attached through
  // the definition part's own relationship.
  const pivotCaches: DefinitionPivotCacheOptions[] = [];
  const pivotCacheIdByPath = new Map<string, number>();
  const definitionCachesByPath = new Map<string, DefinitionPivotCacheOptions>();
  const wbPivotCaches = (xlsx.workbook?.elements ?? []).find(
    (element) => element.name?.replace(/^.*:/, "") === "pivotCaches",
  );
  for (const pc of wbPivotCaches?.elements ?? []) {
    if (pc.name?.replace(/^.*:/, "") !== "pivotCache") continue;
    const cacheId = attr(pc, "cacheId");
    const rId = Object.entries(pc.attributes ?? {}).find(
      ([name]) => name.replace(/^.*:/, "") === "id" && name.includes(":"),
    )?.[1];
    if (cacheId === undefined || rId === undefined) continue;
    const target = readContext.resolveWorksheetRel(xlsx.workbookPath, String(rId));
    if (target) pivotCacheIdByPath.set(target, Number(cacheId));
  }
  for (const definitionPath of xlsx.doc.keys()) {
    if (!/(?:^|\/)pivotCache\/pivotCacheDefinition\d*\.xml$/.test(definitionPath)) continue;
    const definitionEl = xlsx.doc.get(definitionPath);
    if (!definitionEl) continue;
    const definition = parsePivotCacheDefinition(definitionPath, definitionEl);
    const definitionRels = xlsx.doc.get(partPathToRelsPath(definitionPath))?.elements ?? [];
    const rawDefinitionRelationships = definitionRels.filter(
      (rel) => rel.name === "Relationship" && rel.attributes?.["Id"] !== undefined,
    );
    let recordsPath = readContext
      .getWorksheetRelsByType(definitionPath, "/pivotCacheRecords")
      .find(
        (rel) => !definition.recordsRelationshipId || rel.rId === definition.recordsRelationshipId,
      )?.target;
    const recordsEl = recordsPath ? xlsx.doc.get(recordsPath) : undefined;
    const records = recordsEl ? parsePivotCacheRecords(recordsPath!, recordsEl) : undefined;
    let recordsRelationshipSourceId: string | undefined;
    let recordsRelationshipTarget: string | undefined;
    if (recordsPath) {
      const recordsRel = readContext
        .getWorksheetRelsByType(definitionPath, "/pivotCacheRecords")
        .find((rel) => rel.target === recordsPath);
      recordsRelationshipSourceId = recordsRel?.rId;
      recordsRelationshipTarget = rawDefinitionRelationships.find(
        (rel) => rel.attributes?.["Id"] === recordsRelationshipSourceId,
      )?.attributes?.["Target"] as string | undefined;
    }
    const externalRelationships = rawDefinitionRelationships
      .filter((rel) => rel.attributes?.["Id"] !== recordsRelationshipSourceId)
      .map((rel) => ({
        id: String(rel.attributes?.["Id"]),
        type: String(rel.attributes?.["Type"]),
        target: String(rel.attributes?.["Target"]),
        ...(rel.attributes?.["TargetMode"] === "External"
          ? { targetMode: "External" as const }
          : {}),
      }));
    const cacheId = pivotCacheIdByPath.get(definitionPath) ?? pivotCaches.length + 1;
    if (recordsRelationshipSourceId)
      definition.recordsRelationshipSourceId = recordsRelationshipSourceId;
    if (recordsRelationshipTarget) definition.recordsRelationshipTarget = recordsRelationshipTarget;
    definition.relationshipOrder = rawDefinitionRelationships.map((rel) =>
      String(rel.attributes?.["Id"]),
    );
    if (externalRelationships.length > 0) definition.externalRelationships = externalRelationships;
    definitionCachesByPath.set(definitionPath, {
      mode: "definition",
      cacheId,
      definitionPath,
      ...(recordsPath ? { recordsPath } : {}),
      definition,
      ...(records ? { records } : {}),
    });
  }
  const sourceOrderedCachePaths = new Set<string>();
  for (const pc of wbPivotCaches?.elements ?? []) {
    if (pc.name?.replace(/^.*:/, "") !== "pivotCache") continue;
    const cacheId = attr(pc, "cacheId");
    const rId = Object.entries(pc.attributes ?? {}).find(
      ([name]) => name.replace(/^.*:/, "") === "id" && name.includes(":"),
    )?.[1];
    if (cacheId === undefined || rId === undefined) continue;
    const definitionPath = readContext.resolveWorksheetRel(xlsx.workbookPath, String(rId));
    const cache = definitionPath ? definitionCachesByPath.get(definitionPath) : undefined;
    if (!cache || sourceOrderedCachePaths.has(definitionPath!)) continue;
    pivotCaches.push(cache);
    sourceOrderedCachePaths.add(definitionPath!);
  }
  for (const [definitionPath, cache] of definitionCachesByPath) {
    if (!sourceOrderedCachePaths.has(definitionPath)) pivotCaches.push(cache);
  }
  if (pivotCaches.length > 0) opts.pivotCaches = pivotCaches;

  // Parse styles (fonts, fills, borders, cellXfs)
  if (xlsx.styles) {
    const parsedStyles = stylesDesc.parse(xlsx.styles, readContext);

    // Expose styles sections onto the returned opts for round-trip. The six
    // table sections fall back to [] so `undefined` keeps meaning "fresh
    // document" — an adopted table (even empty, e.g. a bare <styleSheet/>) is
    // distinct from the compiler's fresh-file defaults. fonts/fills/borders/
    // cellXfs/numFmts adopt the parsed table wholesale: cells then keep raw
    // style indices (the source's own numbering) instead of resolved
    // definitions — the SDK's Stylesheet model.
    if (parsedStyles.dxfs) opts.dxfs = parsedStyles.dxfs;
    opts.fonts = parsedStyles.fonts ?? [];
    if (parsedStyles.fontsContainer) opts.fontsContainer = parsedStyles.fontsContainer;
    opts.fills = parsedStyles.fills ?? [];
    opts.borders = parsedStyles.borders ?? [];
    opts.cellXfs = parsedStyles.cellXfs ?? [];
    if (parsedStyles.numFmts) opts.numFmts = parsedStyles.numFmts;
    if (parsedStyles.colors) opts.colors = parsedStyles.colors;
    // Optional sections: undefined (absent) stays absent; a present-but-empty
    // section round-trips as an empty container.
    if (parsedStyles.customCellStyles !== undefined)
      opts.cellStyles = parsedStyles.customCellStyles;
    if (parsedStyles.cellStyleXfs !== undefined) opts.cellStyleXfs = parsedStyles.cellStyleXfs;
    if (parsedStyles.styleExtensions) opts.styleExtensions = parsedStyles.styleExtensions;
    if (parsedStyles.tableStylesInfo) opts.tableStyles = parsedStyles.tableStylesInfo;
  }

  // Theme — structured round-trip so a custom source theme survives instead of
  // being replaced by the compiler's fresh default. Parsed under the theme
  // part's own rels scope so a blip fill's r:embed resolves to the theme's
  // image, not to whatever the workbook rels hide under the same rId.
  if (xlsx.theme) {
    const themeEl = xlsx.doc.get(xlsx.theme);
    if (themeEl) {
      const themeOptions = readContext.withPart(xlsx.theme, () =>
        themeDesc.parse(themeEl, readContext),
      );
      if (themeOptions) opts.theme = themeOptions;
    }
  }

  // Parse workbook via descriptor for richer data
  const sheetInfoByPath = new Map<
    string,
    {
      name: string;
      sheetId?: number;
      tabId?: number;
      state?: "visible" | "hidden" | "veryHidden";
    }
  >();
  if (xlsx.workbook) {
    const wbData = workbookDesc.parse(xlsx.workbook, readContext);
    if (wbData.sheets) {
      for (const sheet of wbData.sheets) {
        const target = readContext.resolveWorksheetRel(xlsx.workbookPath, sheet.rId);
        if (target) sheetInfoByPath.set(target, sheet);
      }
      opts.sheetDefinitions = wbData.sheets;
    }

    // Workbook-level properties
    if (wbData.protection) opts.workbookProtection = wbData.protection;
    if (wbData.bookView) opts.bookView = wbData.bookView;
    if (wbData.fileVersion !== undefined) opts.fileVersion = wbData.fileVersion;
    if (wbData.calculation) opts.calculation = wbData.calculation;
    if (wbData.oleSize) opts.oleSize = wbData.oleSize;
    if (wbData.customViews) opts.customWorkbookViews = wbData.customViews;
    if (wbData.fileRecovery) opts.fileRecovery = wbData.fileRecovery;
    if (wbData.functionGroups) opts.functionGroups = wbData.functionGroups;
    if (wbData.webPublishing) opts.webPublishing = wbData.webPublishing;
    if (wbData.fileSharing) opts.fileSharing = wbData.fileSharing;
    if (wbData.properties) opts.properties = wbData.properties;
    if (wbData.conformance) opts.conformance = wbData.conformance;
    if (wbData.webPublishObjects) opts.webPublishObjects = wbData.webPublishObjects;
    if (wbData.definedNames) opts.definedNames = wbData.definedNames;
    if (wbData.absPath !== undefined) opts.absPath = wbData.absPath;
    if (wbData.absPathLegacyPrefix) opts.absPathLegacyPrefix = true;
    if (wbData.legacyChildOrder) opts.legacyChildOrder = true;
    if (wbData.revisionPtr) opts.revisionPtr = wbData.revisionPtr;
    if (wbData.extensions) opts.extensions = wbData.extensions;
  }

  // Parse worksheets using descriptor pipeline. Worksheet parts defer
  // sheetData — the row scanner walks the captured inner XML, skipping the
  // per-cell Element tree (the dominant allocation cost on large sheets).
  const worksheets: WorksheetOptions[] = [];
  const chartExternalLinkPaths = new Set<string>();
  // Drawing/chart parts the canonical model actually absorbed — anything the
  // model never reached (broken worksheet rels, missing anchors) stays a
  // passthrough part instead of being dropped as a "rebuilt" path.
  const absorbedDrawingParts = new Set<string>();
  const absorbedChartParts = new Set<string>();
  const absorbedCommentsParts = new Set<string>();
  const absorbedControlParts = new Set<string>();
  for (const wsPath of xlsx.worksheets) {
    const wsEl = xlsx.doc.get(wsPath, WORKSHEET_PARSE_OPTIONS);
    if (!wsEl) continue;

    const wsOpts = readContext.withPart(wsPath, () => worksheetDesc.parse(wsEl, readContext));
    wsOpts.sourcePath = wsPath;
    const sheetInfo = sheetInfoByPath.get(wsPath);
    if (sheetInfo) {
      wsOpts.name = sheetInfo.name;
      wsOpts.sheetId = sheetInfo.sheetId ?? sheetInfo.tabId;
      if (sheetInfo.state) wsOpts.state = sheetInfo.state;
    }

    // ── Resolve sub-parts via worksheet relationships ──

    // Comments
    const commentRels = readContext.getWorksheetRelsByType(wsPath, "/comments");
    for (const cr of commentRels) {
      const commentEl = xlsx.doc.get(cr.target);
      if (!commentEl) continue;
      const commentData = commentsDesc.parse(commentEl, readContext);
      if (commentData.comments) {
        wsOpts.comments = commentData.comments;
        absorbedCommentsParts.add(cr.target);
        break; // one comments file per worksheet
      }
    }

    // Note anchors (vmlDrawing) — merge per-note placement into the comments
    // so custom position/size/visibility survive the round-trip.
    const vmlRels = readContext.getWorksheetRelsByType(wsPath, "/vmlDrawing");
    for (const vr of vmlRels) {
      const vmlEl = xlsx.doc.get(vr.target);
      if (!vmlEl) continue;
      const vml = vmlNotesDesc.parse(vmlEl, readContext);
      mergeNoteAnchors(wsOpts.comments, vml.anchors);
      if (vml.layout) wsOpts.commentsVmlLayout = vml.layout;
      if (vml.shapeType) wsOpts.commentsVmlShapeType = vml.shapeType;
      const source = xlsx.doc.getRaw(vr.target);
      if (source) {
        wsOpts.commentsVmlBom =
          source.length >= 3 && source[0] === 0xef && source[1] === 0xbb && source[2] === 0xbf;
        wsOpts.commentsVmlSource = new TextDecoder().decode(source);
      }
      break; // one vmlDrawing per worksheet
    }

    // Form-control properties and ActiveX metadata live in per-control parts.
    for (const control of wsOpts.controls ?? []) {
      const target = readContext.resolveWorksheetRel(wsPath, control.rId!);
      if (target === undefined) {
        throw new ControlPropertiesParseError(
          "xl/worksheets",
          wsPath,
          control.rId!,
          "unresolved control relationship",
        );
      }
      if (control.kind === "form") {
        const el = xlsx.doc.get(target);
        if (!el) {
          throw new ControlPropertiesParseError(
            "xl/ctrlProps",
            target,
            "formControlPr",
            "missing part",
          );
        }
        control.formControlProperties = controlPropertiesDesc.parse(el, readContext);
      } else {
        const el = xlsx.doc.get(target);
        if (!el) {
          throw new ActiveXControlParseError("xl/activeX", target, "ax:ocx", "missing part");
        }
        control.activeXControl = activeXControlDesc.parse(el, readContext);
        const relsEl = xlsx.doc.get(partPathToRelsPath(target));
        for (const child of relsEl?.elements ?? []) {
          if (child.name !== "Relationship") {
            throw new ActiveXControlParseError(
              target,
              partPathToRelsPath(target),
              String(child.name ?? ""),
              "unsupported relationship element",
            );
          }
          const type = attr(child, "Type");
          const relTarget = attr(child, "Target");
          const rId = attr(child, "Id");
          if (!type || !relTarget || !rId) {
            throw new ActiveXControlParseError(
              target,
              partPathToRelsPath(target),
              "Relationship",
              "missing required relationship attribute",
            );
          }
          if (!type.endsWith("/activeXControlBinary")) {
            throw new ActiveXControlParseError(
              target,
              partPathToRelsPath(target),
              type,
              "unsupported relationship type",
            );
          }
          control.activeXControl.binaryPath = resolveRelationshipTarget(target, relTarget);
        }
      }
      absorbedControlParts.add(target);
      absorbedControlParts.add(partPathToRelsPath(target));
    }

    // Drawings (images + charts)
    const drawingRels = readContext.getWorksheetRelsByType(wsPath, "/drawing");
    for (const dr of drawingRels) {
      const drawingEl = xlsx.doc.get(dr.target);
      if (!drawingEl) continue;
      wsOpts.drawingPresent = true;
      // cNvPr hyperlinks (a:hlinkClick) resolve through the drawing part's own
      // rels: internal targets resolve against the part path, External ones
      // (absolute URLs) stay verbatim. Fall back to the workbook context.
      const drawingRelById = new Map<
        string,
        { target: string; mode?: string; relationshipType: string }
      >();
      const drawingRelsEl = xlsx.doc.get(partPathToRelsPath(dr.target));
      for (const rel of drawingRelsEl?.elements ?? []) {
        if (rel.name !== "Relationship") continue;
        const id = rel.attributes?.["Id"];
        const target = rel.attributes?.["Target"];
        if (id === undefined || target === undefined) continue;
        const mode =
          rel.attributes?.["TargetMode"] !== undefined
            ? String(rel.attributes["TargetMode"])
            : undefined;
        const relationshipType = rel.attributes?.["Type"];
        if (relationshipType === undefined) continue;
        drawingRelById.set(String(id), {
          target: String(target),
          mode,
          relationshipType: String(relationshipType),
        });
      }
      const drawingCtx: ReadContext = {
        resolveRelationship: (rid) => {
          const rel = drawingRelById.get(rid);
          if (!rel) return readContext.resolveRelationship(rid);
          return rel.mode === "External"
            ? rel.target
            : resolveRelationshipTarget(dr.target, rel.target);
        },
        getPart: (path) => readContext.getPart(path),
        getRaw: (path) => readContext.getRaw(path),
      };
      const drawingData = drawingDesc.parse(drawingEl, drawingCtx);
      absorbedDrawingParts.add(dr.target);
      // drawingDesc.parse yields CT-layer anchors (rId-anchored DrawingImage/
      // DrawingChart); bridge them to the user-layer shapes the compiler
      // consumes: image bytes are read back through the drawing's image
      // relationships, chart parts through the core chartSpace descriptor.
      if (drawingData.images) {
        const images: PictureOptions[] = [];
        for (const image of drawingData.images) {
          // Linked source (a:blip @r:link): only an External relationship
          // carries a usable URL; internal link targets have no media part.
          const linkRel = image.linkRId ? drawingRelById.get(image.linkRId) : undefined;
          const sourceUrl = linkRel?.mode === "External" ? linkRel.target : undefined;
          const mediaPath = image.rId
            ? readContext.resolveWorksheetRel(dr.target, image.rId)
            : undefined;
          const raw = mediaPath ? xlsx.doc.getRaw(mediaPath) : undefined;
          const ext = mediaPath?.split(".").pop();
          if (
            !raw ||
            (ext !== "png" && ext !== "jpeg" && ext !== "jpg" && ext !== "wmf" && ext !== "emf")
          ) {
            // Linked-only picture (no bytes in the package): keep the URL,
            // derive the type token from it (png fallback for extension-less).
            if (sourceUrl !== undefined) {
              const linkExt = sourceUrl.split(".").pop()?.toLowerCase() ?? "";
              images.push({
                type:
                  linkExt === "wmf"
                    ? "wmf"
                    : linkExt === "emf"
                      ? "emf"
                      : linkExt === "jpg" || linkExt === "jpeg"
                        ? "jpg"
                        : "png",
                sourceUrl,
                ...pickAnchorOptions(image),
                name: image.name,
                description: image.description,
                title: image.title,
                hidden: image.hidden,
                ...(image.creationId !== undefined ? { creationId: image.creationId } : {}),
                ...(image.ext !== undefined ? { ext: image.ext } : {}),
                ...(image.properties ? { properties: image.properties } : {}),
                ...(image.blackWhiteMode ? { blackWhiteMode: image.blackWhiteMode } : {}),
                ...(image.compression !== undefined ? { compression: image.compression } : {}),
                ...(image.sourceRectangle ? { sourceRectangle: image.sourceRectangle } : {}),
                ...(image.preferRelativeResize !== undefined
                  ? { preferRelativeResize: image.preferRelativeResize }
                  : {}),
                ...(image.blipEffects ? { blipEffects: image.blipEffects } : {}),
                ...(image.useLocalDpi !== undefined ? { useLocalDpi: image.useLocalDpi } : {}),
                ...(image.blipExt !== undefined ? { blipExt: image.blipExt } : {}),
                ...(image.locking ? { locking: image.locking } : {}),
                ...(image.hyperlink ? { hyperlink: image.hyperlink } : {}),
                ...(image.zOrder !== undefined ? { zOrder: image.zOrder } : {}),
                ...(image.shapeId !== undefined ? { shapeId: image.shapeId } : {}),
              });
            }
            continue;
          }
          // WMF/EMF clip-art images round-trip like raster ones (the media
          // store keeps their bytes and extension verbatim).
          const type =
            ext === "png" ? "png" : ext === "wmf" ? "wmf" : ext === "emf" ? "emf" : "jpg";
          images.push({
            data: raw,
            type,
            ...(mediaPath ? { sourcePath: mediaPath } : {}),
            ...(sourceUrl !== undefined ? { sourceUrl } : {}),
            ...pickAnchorOptions(image),
            name: image.name,
            description: image.description,
            title: image.title,
            hidden: image.hidden,
            ...(image.creationId !== undefined ? { creationId: image.creationId } : {}),
            ...(image.ext !== undefined ? { ext: image.ext } : {}),
            ...(image.properties ? { properties: image.properties } : {}),
            ...(image.blackWhiteMode ? { blackWhiteMode: image.blackWhiteMode } : {}),
            ...(image.compression !== undefined ? { compression: image.compression } : {}),
            ...(image.sourceRectangle ? { sourceRectangle: image.sourceRectangle } : {}),
            ...(image.preferRelativeResize !== undefined
              ? { preferRelativeResize: image.preferRelativeResize }
              : {}),
            ...(image.blipEffects ? { blipEffects: image.blipEffects } : {}),
            ...(image.useLocalDpi !== undefined ? { useLocalDpi: image.useLocalDpi } : {}),
            ...(image.blipExt !== undefined ? { blipExt: image.blipExt } : {}),
            ...(image.locking ? { locking: image.locking } : {}),
            ...(image.hyperlink ? { hyperlink: image.hyperlink } : {}),
            ...(image.zOrder !== undefined ? { zOrder: image.zOrder } : {}),
            ...(image.shapeId !== undefined ? { shapeId: image.shapeId } : {}),
          });
        }
        if (images.length > 0) wsOpts.images = images;
      }
      if (drawingData.charts) {
        const charts: WorksheetChartOptions[] = [];
        for (const anchor of drawingData.charts) {
          const chartPath = readContext.resolveWorksheetRel(dr.target, anchor.rId);
          if (!chartPath) continue;
          const chartEl = xlsx.doc.get(chartPath);
          if (!chartEl) continue;
          const chartSpace = readContext.withPart(chartPath, () =>
            chartSpaceDesc.parse(chartEl, readContext),
          );
          absorbedChartParts.add(chartPath);
          readChartUserShapes(chartPath, chartSpace, readContext, xlsx.doc);
          const chartExternalLink = readChartExternalLink(
            chartPath,
            chartSpace.externalData,
            readContext,
            xlsx.doc,
            chartExternalLinkPaths,
          );
          // cNvPr @title stays unbridged (same rule as the compiler leg):
          // WorksheetChartOptions.title is the chart title, not the frame's.
          const chartCnvPr = pickNonVisualDrawingProperties(anchor);
          delete chartCnvPr.title;
          charts.push({
            ...chartSpace,
            ...pickAnchorOptions(anchor),
            ...chartCnvPr,
            ...(anchor.frameExtentCx !== undefined ? { frameExtentCx: anchor.frameExtentCx } : {}),
            ...(anchor.frameExtentCy !== undefined ? { frameExtentCy: anchor.frameExtentCy } : {}),
            ...(chartPath ? { sourcePath: chartPath } : {}),
            ...chartExternalLink,
            ...(anchor.frameLocks ? { frameLocks: anchor.frameLocks } : {}),
            ...(anchor.macro !== undefined ? { macro: anchor.macro } : {}),
            ...(anchor.fPublished !== undefined ? { fPublished: anchor.fPublished } : {}),
            ...(anchor.hyperlink ? { hyperlink: anchor.hyperlink } : {}),
            ...(anchor.zOrder !== undefined ? { zOrder: anchor.zOrder } : {}),
            ...(anchor.shapeId !== undefined ? { shapeId: anchor.shapeId } : {}),
          });
        }
        if (charts.length > 0) wsOpts.charts = charts;
      }
      if (drawingData.smartArts) {
        const smartArts: WorksheetSmartArtOptions[] = [];
        for (const anchor of drawingData.smartArts) {
          // The diagram parts themselves pass through verbatim; carry their
          // package-absolute paths so the compiler can re-wire the drawing rels.
          const dataPath = readContext.resolveWorksheetRel(dr.target, anchor.dataRId);
          const layoutPath = readContext.resolveWorksheetRel(dr.target, anchor.layoutRId);
          const quickStylePath = readContext.resolveWorksheetRel(dr.target, anchor.quickStyleRId);
          const colorsPath = readContext.resolveWorksheetRel(dr.target, anchor.colorsRId);
          if (!dataPath || !layoutPath || !quickStylePath || !colorsPath) continue;
          smartArts.push({
            ...pickAnchorOptions(anchor),
            ...pickNonVisualDrawingProperties(anchor),
            dataPath,
            layoutPath,
            quickStylePath,
            colorsPath,
            ...(anchor.frameLocks ? { frameLocks: anchor.frameLocks } : {}),
            ...(anchor.macro !== undefined ? { macro: anchor.macro } : {}),
            ...(anchor.zOrder !== undefined ? { zOrder: anchor.zOrder } : {}),
            ...(anchor.shapeId !== undefined ? { shapeId: anchor.shapeId } : {}),
          });
        }
        if (smartArts.length > 0) wsOpts.smartArts = smartArts;
      }
      if (drawingData.webExtensions) {
        const webExtensions: WorksheetWebExtensionOptions[] = [];
        for (const anchor of drawingData.webExtensions) {
          const sourcePath = readContext.resolveWorksheetRel(dr.target, anchor.rId);
          if (!sourcePath) continue;
          const snapshotSourcePath = anchor.fallback?.rId
            ? readContext.resolveWorksheetRel(dr.target, anchor.fallback.rId)
            : undefined;
          webExtensions.push({
            ...pickAnchorOptions(anchor),
            ...pickNonVisualDrawingProperties(anchor),
            sourcePath,
            ...(anchor.elementName ? { elementName: anchor.elementName } : {}),
            ...(snapshotSourcePath ? { snapshotSourcePath } : {}),
            ...(anchor.fallback ? { fallback: anchor.fallback } : {}),
            ...(anchor.frameLocks ? { frameLocks: anchor.frameLocks } : {}),
            ...(anchor.macro !== undefined ? { macro: anchor.macro } : {}),
            ...(anchor.zOrder !== undefined ? { zOrder: anchor.zOrder } : {}),
            ...(anchor.shapeId !== undefined ? { shapeId: anchor.shapeId } : {}),
          });
        }
        if (webExtensions.length > 0) wsOpts.webExtensions = webExtensions;
      }
      // Shapes/connectors/groups pass through unchanged (no media bridge).
      if (drawingData.shapes) wsOpts.shapes = drawingData.shapes;
      if (drawingData.connectors) wsOpts.connectors = drawingData.connectors;
      if (drawingData.groups) wsOpts.groups = drawingData.groups;
      if (drawingData.contentParts) {
        const contentParts: WorksheetContentPartOptions[] = [];
        for (const contentPart of drawingData.contentParts) {
          const relationship = drawingRelById.get(contentPart.rId);
          if (!relationship) continue;
          contentParts.push({
            ...contentPart,
            relationshipType: relationship.relationshipType,
            relationshipTarget: relationship.target,
            sourcePath: resolveRelationshipTarget(dr.target, relationship.target),
          });
        }
        if (contentParts.length > 0) wsOpts.contentParts = contentParts;
      }
      break;
    }

    // Background picture — the worksheet-level image relationship backs the
    // <picture r:id/> element (drawing images live in the drawing part's own
    // rels, so worksheet-level image rels are backgrounds only).
    const bgRels = readContext.getWorksheetRelsByType(wsPath, "/image");
    for (const bg of bgRels) {
      const raw = xlsx.doc.getRaw(bg.target);
      const ext = bg.target.split(".").pop();
      if (!raw || (ext !== "png" && ext !== "jpeg" && ext !== "jpg")) continue;
      wsOpts.backgroundImage = { data: raw, type: ext === "png" ? "png" : "jpg" };
      break; // one background picture per worksheet
    }

    // Tables
    const tableRels = readContext.getWorksheetRelsByType(wsPath, "/table");
    if (tableRels.length > 0) {
      const tables: TableOptions[] = [];
      for (const tr of tableRels) {
        const tableEl = xlsx.doc.get(tr.target);
        if (!tableEl) continue;
        // The /table relationship also targets XML Map parts (singleXmlCells)
        // whose root is not a table — skip anything else instead of feeding
        // an unmodeled part through the table descriptor.
        if (tableEl.name !== "table") continue;
        const tableData = tableDesc.parse(tableEl, readContext);
        tables.push(tableData);
      }
      if (tables.length > 0) wsOpts.tables = tables;
    }
    if (!wsOpts.tables?.length && findChild(wsEl, "tableParts")) {
      wsOpts.preserveEmptyTableParts = true;
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

    // Pivot tables remain complete CT definitions; each worksheet relationship
    // identifies its part and its cache relationship identifies the cacheId.
    const pivotTableRelTargets = readContext
      .getWorksheetRelsByType(wsPath, "/pivotTable")
      .map((rel) => rel.target);
    for (const pivotTablePath of pivotTableRelTargets) {
      const pivotTableEl = xlsx.doc.get(pivotTablePath);
      if (!pivotTableEl) continue;
      const definition = parsePivotTableDefinition(pivotTablePath, pivotTableEl);
      const cacheRel = readContext
        .getWorksheetRelsByType(pivotTablePath, "/pivotCacheDefinition")
        .find((rel) => pivotCacheIdByPath.has(rel.target));
      const cachePath = cacheRel?.target;
      const cacheId = cachePath ? pivotCacheIdByPath.get(cachePath) : undefined;
      if (cacheId === undefined) continue;
      const rawCacheRel = (xlsx.doc.get(partPathToRelsPath(pivotTablePath))?.elements ?? []).find(
        (rel) => rel.name === "Relationship" && rel.attributes?.["Id"] === cacheRel?.rId,
      );
      (wsOpts.pivotTables ??= []).push({
        mode: "definition",
        cacheId,
        definition: {
          ...definition,
          cacheRelationshipId: cacheRel?.rId,
          cacheRelationshipTarget: rawCacheRel?.attributes?.["Target"] as string | undefined,
          cacheRelationshipType: rawCacheRel?.attributes?.["Type"] as string | undefined,
        },
      });
    }

    // Resolve external hyperlink URLs
    const hyperlinks = wsOpts.hyperlinks;
    if (hyperlinks) {
      for (const hl of hyperlinks) {
        if (hl.url !== undefined) {
          const resolved = readContext.resolveWorksheetRel(wsPath, hl.url);
          if (resolved) hl.url = resolved;
        }
      }
    }

    worksheets.push(wsOpts as WorksheetOptions);
  }

  opts.worksheets = worksheets;

  // Chartsheets — parse chartsheet parts. Sort numerically: doc.keys() yields
  // ZIP entry order, while sheetNames indexes rely on sheetN.xml numbering.
  const chartsheetPaths = sortByNumber(
    xlsx.doc.keys("xl/chartsheets/").filter((k) => k.endsWith(".xml")),
  );
  if (chartsheetPaths.length > 0) {
    const chartsheets: ChartsheetOptions[] = [];
    for (const csPath of chartsheetPaths) {
      const csEl = xlsx.doc.get(csPath);
      if (!csEl) continue;
      const csData = chartsheetDesc.parse(csEl, readContext);
      const sheetInfo = sheetInfoByPath.get(csPath);
      if (sheetInfo) {
        csData.name = sheetInfo.name;
        csData.sheetId = sheetInfo.sheetId;
        if (sheetInfo.state) csData.state = sheetInfo.state;
      }
      // The chart itself lives in a drawing part — bridge it back through the
      // core chartSpace descriptor into the simplified chartsheet chart shape.
      const csDrawingRels = readContext.getWorksheetRelsByType(csPath, "/drawing");
      outer: for (const dr of csDrawingRels) {
        const drawingPath = dr.target;
        const drawingEl = xlsx.doc.get(drawingPath);
        if (!drawingEl) continue;
        csData.sourceDrawingPath = drawingPath;
        csData.sourceDrawingRelationshipId = dr.rId;
        const drawingData = drawingDesc.parse(drawingEl, readContext);
        absorbedDrawingParts.add(drawingPath);
        for (const anchor of drawingData.charts ?? []) {
          const chartPath = readContext.resolveWorksheetRel(drawingPath, anchor.rId);
          const chartEl = chartPath ? xlsx.doc.get(chartPath) : undefined;
          if (!chartPath || !chartEl) continue;
          // Full chartSpace passthrough — the simplified type/title/series
          // projection dropped chartSpace-level fidelity (c:lang, c:date1904,
          // axis/plot formatting, …) on round-trip.
          csData.chart = chartSpaceDesc.parse(chartEl, readContext);
          csData.sourceChartPath = chartPath;
          absorbedChartParts.add(chartPath);
          readChartUserShapes(chartPath, csData.chart, readContext, xlsx.doc);
          Object.assign(
            csData,
            readChartExternalLink(
              chartPath,
              csData.chart.externalData,
              readContext,
              xlsx.doc,
              chartExternalLinkPaths,
            ),
          );
          if (anchor.macro !== undefined) csData.macro = anchor.macro;
          if (anchor.frameLocks) csData.frameLocks = anchor.frameLocks;
          // Anchor geometry is the rendered chart size on the sheet — Excel
          // keeps it verbatim on save, so round-trip must not fall back to the
          // full-page default. The drawing fields accept UniversalMeasure;
          // chartsheet anchors are plain EMU numbers.
          if (anchor.absoluteX !== undefined) csData.absoluteX = convertToEmu(anchor.absoluteX);
          if (anchor.absoluteY !== undefined) csData.absoluteY = convertToEmu(anchor.absoluteY);
          if (anchor.extentCx !== undefined) csData.extentCx = convertToEmu(anchor.extentCx);
          if (anchor.extentCy !== undefined) csData.extentCy = convertToEmu(anchor.extentCy);
          if (anchor.shapeId !== undefined) csData.shapeId = anchor.shapeId;
          const { name, ...chartIdentity } = pickNonVisualDrawingProperties(anchor);
          if (name !== undefined) csData.chartName = name;
          Object.assign(csData, chartIdentity);
          break outer;
        }
      }
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

  const pivotTablePaths = xlsx.worksheets.flatMap((worksheetPath) =>
    readContext.getWorksheetRelsByType(worksheetPath, "/pivotTable").map((rel) => rel.target),
  );

  // Calculation chain
  const calcChainEl = xlsx.doc.get("xl/calcChain.xml");
  if (calcChainEl) {
    const calcData = calcChainDesc.parse(calcChainEl, readContext);
    if (calcData.cells) opts.calcChain = calcData.cells;
  } else {
    opts.calcChain = false;
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

  // Volatile function types (xl/volTypes.xml)
  const volTypesPath = xlsx.doc.has("xl/volTypes.xml")
    ? "volTypes.xml"
    : xlsx.doc.has("xl/volatileDependencies.xml")
      ? "volatileDependencies.xml"
      : undefined;
  const volTypesEl = volTypesPath ? xlsx.doc.get(`xl/${volTypesPath}`) : undefined;
  if (volTypesEl) {
    const volTypes = parseVolTypesEl(volTypesEl);
    opts.volTypes = volTypes;
    opts.volTypesPath = volTypesPath;
    if (volTypes.length > 0) {
      opts.volTypesCount = attrNum(volTypesEl, "count");
    }
  }

  // External links
  const extLinkPaths = xlsx.doc
    .keys("xl/externalLinks/")
    .filter((k) => k.endsWith(".xml") && !chartExternalLinkPaths.has(k));
  if (extLinkPaths.length > 0) {
    const externalLinks: ExternalLinkOptions[] = [];
    for (const elPath of extLinkPaths) {
      const elEl = xlsx.doc.get(elPath);
      if (!elEl) continue;
      const elData = externalLinkDesc.parse(elEl, readContext);
      elData.sourcePath = elPath;

      // Resolve the external book target from the sibling rels file
      // (xl/externalLinks/_rels/externalLinkN.xml.rels), which compiler.ts writes.
      if (elData.externalBook) {
        const relsPath = partPathToRelsPath(elPath);
        const relsEl = xlsx.doc.get(relsPath);
        if (relsEl) {
          for (const child of relsEl.elements ?? []) {
            if (child.name !== "Relationship") continue;
            const type = attr(child, "Type") ?? "";
            if (!type.toLowerCase().includes("externallinkpath")) continue;
            const target = attr(child, "Target");
            if (target) {
              elData.externalBook.target = target;
              elData.externalBook.targetTypeRaw = type;
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
  const wbRelsEl2 = xlsx.doc.get(partPathToRelsPath(xlsx.workbookPath));
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
    const revisionHeadersPath = (target: string): string => {
      const resolved = resolveRelationshipTarget(xlsx.workbookPath, target);
      return resolved.startsWith("xl/") ? resolved.slice(3) : resolved;
    };
    const headersRelativePath = revisionHeadersPath(revHeadersTarget);
    const headersPath = `xl/${headersRelativePath}`;
    const headersEl = xlsx.doc.get(headersPath);
    if (headersEl) {
      const headers = revisionHeadersDesc.parse(headersEl, readContext);
      const logsByRid = new Map<string, RevisionLogOptions>();
      const revHeadersRelsEl = xlsx.doc.get(partPathToRelsPath(headersPath));
      if (revHeadersRelsEl) {
        for (const child of revHeadersRelsEl.elements ?? []) {
          if (child.name !== "Relationship") continue;
          if (!(attr(child, "Type") ?? "").includes("/revisionLog")) continue;
          const t = attr(child, "Target");
          const rId = attr(child, "Id");
          if (!t || !rId) continue;
          const logPath = resolveRelationshipTarget(headersPath, t);
          const logEl = xlsx.doc.get(logPath);
          if (!logEl) continue;
          logsByRid.set(rId, {
            ...revisionLogDesc.parse(logEl, readContext),
            path: logPath.startsWith("xl/") ? logPath.slice(3) : logPath,
            relationshipTarget: t,
          });
        }
      }
      const logs = headers.headers
        .map((header) => logsByRid.get(header.rId))
        .filter((log): log is RevisionLogOptions => log !== undefined);
      const revisionLog: SharedWorkbookOptions = {
        headers,
        logs,
        headersPath: headersRelativePath,
        ...(usersTarget ? { usersPath: revisionHeadersPath(usersTarget) } : {}),
      };
      if (usersTarget) {
        const usersEl = xlsx.doc.get(`xl/${revisionLog.usersPath}`);
        if (usersEl) {
          const users = usersDesc.parse(usersEl, readContext);
          if (users.users) revisionLog.users = users;
        }
      }
      opts.revisionLog = revisionLog;
    }
  }

  // Package-wide passthrough (SDK ExtendedPart analogue): every part the model
  // did NOT absorb is carried verbatim instead of dropped. Only parts the
  // compiler always re-emits are excluded — model-driven parts (themes,
  // sharedStrings, drawings, VML, external links) pass through and yield to
  // the compiler's own output at the same path by assembly order.
  const rebuilt: string[] = [
    xlsx.workbookPath,
    partPathToRelsPath(xlsx.workbookPath),
    ...(xlsx.coreProps ? [xlsx.coreProps] : []),
    ...(xlsx.appProps ? [xlsx.appProps] : []),
    ...(xlsx.customProps ? [xlsx.customProps] : []),
    ...xlsx.worksheets,
    ...xlsx.worksheets.map((path) => partPathToRelsPath(path)),
    ...chartsheetPaths,
    ...dialogsheetPaths,
    ...extLinkPaths,
    ...extLinkPaths.map((path) => partPathToRelsPath(path)),
    ...(xlsx.styles ? [xlsx.stylesPath, partPathToRelsPath(xlsx.stylesPath)] : []),
    ...(xlsx.theme ? [xlsx.theme, partPathToRelsPath(xlsx.theme)] : []),
    ...(xlsx.sharedStrings ? [xlsx.sharedStringsPath] : []),
    ...(calcChainEl ? ["xl/calcChain.xml"] : []),
    ...(connectionsEl ? ["xl/connections.xml"] : []),
    ...(volTypesPath === "volatileDependencies.xml" ? ["xl/volatileDependencies.xml"] : []),
    ...xlsx.partRefs.charts.filter((path) => absorbedChartParts.has(path)),
    ...xlsx.partRefs.charts
      .filter((path) => absorbedChartParts.has(path))
      .map((path) => partPathToRelsPath(path)),
    ...[...chartExternalLinkPaths].flatMap((path) => [path, partPathToRelsPath(path)]),
    ...sortByNumber(
      xlsx.doc
        .keys("xl/comments")
        .filter((path) => /^xl\/comments\d+\.xml$/i.test(path) && absorbedCommentsParts.has(path)),
    ).flatMap((path) => [path, partPathToRelsPath(path)]),
    ...xlsx.doc
      .keys("xl/ctrlProps/")
      .filter((path) => path.endsWith(".xml") && absorbedControlParts.has(path)),
    ...xlsx.doc.keys("xl/activeX/").filter((path) => absorbedControlParts.has(path)),
    ...xlsx.worksheets.flatMap((worksheetPath) => {
      if (readContext.getWorksheetRelsByType(worksheetPath, "/comments").length === 0) return [];
      return readContext
        .getWorksheetRelsByType(worksheetPath, "/vmlDrawing")
        .map((relationship) => relationship.target);
    }),
    ...(metadataEl ? ["xl/metadata.xml"] : []),
    ...chartsheetPaths.map((path) => partPathToRelsPath(path)),
    ...xlsx.doc
      .keys("xl/tables/")
      .filter((path) => path.endsWith(".xml") && xlsx.doc.get(path)?.name === "table"),
    ...xlsx.partRefs.drawings
      .filter((path) => absorbedDrawingParts.has(path))
      .flatMap((path) => [path, partPathToRelsPath(path)]),
    ...pivotCaches.flatMap((cache) => [
      cache.definitionPath,
      ...(cache.recordsPath ? [cache.recordsPath] : []),
      partPathToRelsPath(cache.definitionPath),
    ]),
    ...pivotTablePaths,
    ...pivotTablePaths.map((pivotTablePath) => partPathToRelsPath(pivotTablePath)),
    ...(opts.revisionLog
      ? [
          `xl/${opts.revisionLog.headersPath}`,
          partPathToRelsPath(`xl/${opts.revisionLog.headersPath}`),
          ...opts.revisionLog.logs.flatMap((log) => (log.path ? [`xl/${log.path}`] : [])),
          ...(opts.revisionLog.users ? [`xl/${opts.revisionLog.usersPath}`] : []),
        ]
      : []),
  ];
  const { parts: passthroughParts, relationships: passthroughRels } = collectPassthroughParts(
    xlsx.doc,
    rebuilt,
    // Stage-0 rawParts policy: XML parts not listed here are flagged by the
    // audit as modeled-XML absorption gaps (strict policy rejects them).
    opaquePassthroughPolicy("xlsx"),
  );
  passthroughParts.push(
    ...xlsx.doc
      .keys()
      .filter((path) => /^xl\/(?:pivotCache|pivotTables)(?:\/_rels)?\/$/i.test(path))
      .map((path) => ({ path, data: new Uint8Array(0) })),
  );
  if (passthroughParts.length > 0) opts.rawParts = passthroughParts;
  if (passthroughRels.length > 0) opts.passthroughRelationships = passthroughRels;
  if (xlsx.packageRelationshipNamespace)
    opts.relationshipNamespace = xlsx.packageRelationshipNamespace;
  if (xlsx.workbookPath !== "xl/workbook.xml") opts.workbookPath = xlsx.workbookPath;
  if (xlsx.stylesPath !== "xl/styles.xml") opts.stylesPath = xlsx.stylesPath;
  if (xlsx.sharedStringsPath !== "xl/sharedStrings.xml")
    opts.sharedStringsPath = xlsx.sharedStringsPath;
  if (xlsx.theme && xlsx.theme !== "xl/theme/theme1.xml") opts.themePath = xlsx.theme;

  // Source content-type declarations — the compiler keeps them as the base
  // table so round-trip preserves the Default/Override split as written.
  const sourceContentTypes = xlsx.doc.get("[Content_Types].xml");
  if (sourceContentTypes) {
    const ct = contentTypesDesc.parse(sourceContentTypes, {} as ReadContext);
    if (ct) opts.contentTypes = { ...ct, preserveSourceDeclarations: true };
  }

  return opts as WorkbookOptions;
}
