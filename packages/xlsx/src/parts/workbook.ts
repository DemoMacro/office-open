/**
 * Workbook types and descriptor for SpreadsheetML documents.
 *
 * @module
 */

import { derivePasswordHash } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { findChild, attr, attrNum, escapeXml } from "@office-open/xml";

export interface SheetDefinition {
  name: string;
  sheetId: number;
  rId: string;
  state?: "visible" | "hidden" | "veryHidden";
}

export interface PivotCacheReference {
  cacheId: number;
  rId: string;
}

export interface TablePartReference {
  rId: string;
}

/** Custom workbook view for storing display preferences. */
export interface CustomWorkbookViewOptions {
  /** View name */
  name: string;
  /** GUID (e.g. "{00000000-0000-0000-0000-000000000000}") */
  guid: string;
  /** Window width in twips */
  windowWidth: number;
  /** Window height in twips */
  windowHeight: number;
  /** Active sheet ID (1-based sheetId) */
  activeSheetId: number;
  /** X position of the window */
  xWindow?: number;
  /** Y position of the window */
  yWindow?: number;
  /** Show formula bar (default true) */
  showFormulaBar?: boolean;
  /** Show status bar (default true) */
  showStatusbar?: boolean;
  /** Show horizontal scroll (default true) */
  showHorizontalScroll?: boolean;
  /** Show vertical scroll (default true) */
  showVerticalScroll?: boolean;
  /** Show sheet tabs (default true) */
  showSheetTabs?: boolean;
  /** Tab ratio (default 600) */
  tabRatio?: number;
  /** Include hidden rows/columns (default true) */
  includeHiddenRowCol?: boolean;
  /** Include print settings (default true) */
  includePrintSettings?: boolean;
  /** Personal view (default false) */
  personalView?: boolean;
  /** Maximized (default false) */
  maximized?: boolean;
  /** Minimized (default false) */
  minimized?: boolean;
  /** Auto update (CT_CustomWorkbookView @autoUpdate) */
  autoUpdate?: boolean;
  /** Merge interval (CT_CustomWorkbookView @mergeInterval) */
  mergeInterval?: number;
  /** Changes saved in window (CT_CustomWorkbookView @changesSavedWin) */
  changesSavedWin?: boolean;
  /** Only sync (CT_CustomWorkbookView @onlySync) */
  onlySync?: boolean;
  /** Show comments (CT_CustomWorkbookView @showComments) */
  showComments?: string;
}

export interface WorkbookProtectionOptions {
  /** Lock workbook structure (add/delete/rename/move sheets) */
  lockStructure?: boolean;
  /** Lock workbook windows */
  lockWindows?: boolean;
  /** Lock revisions */
  lockRevision?: boolean;
  /** Plain-text password — legacy Excel hash computed automatically */
  workbookPassword?: string;
  /** Modern encryption: algorithm name (e.g. "SHA-512") */
  workbookAlgorithmName?: string;
  /** Modern encryption: base64-encoded hash value */
  workbookHashValue?: string;
  /** Modern encryption: base64-encoded salt value */
  workbookSaltValue?: string;
  /** Modern encryption: spin count */
  workbookSpinCount?: number;
  /** Revisions password (legacy) */
  revisionsPassword?: string;
  /** Revisions modern encryption: algorithm name */
  revisionsAlgorithmName?: string;
  /** Revisions modern encryption: base64-encoded hash value */
  revisionsHashValue?: string;
  /** Revisions modern encryption: base64-encoded salt value */
  revisionsSaltValue?: string;
  /** Revisions modern encryption: spin count */
  revisionsSpinCount?: number;
  /** Workbook password character set (CT_WorkbookProtection @workbookPasswordCharacterSet) */
  workbookPasswordCharacterSet?: string;
  /** Revisions password character set (CT_WorkbookProtection @revisionsPasswordCharacterSet) */
  revisionsPasswordCharacterSet?: string;
}

/** Workbook conformance level (CT_Workbook @conformance) */
export type WorkbookConformance = "strict" | "transitional";

/** File recovery properties (CT_FileRecoveryPr) */
export interface FileRecoveryPrOptions {
  /** Enable auto-recover (default true) */
  autoRecover?: boolean;
  /** Crash save (default false) */
  crashSave?: boolean;
  /** Data extract load (default false) */
  dataExtractLoad?: boolean;
  /** Repair load (default false) */
  repairLoad?: boolean;
}

/** Web publishing properties (CT_WebPublishing) */
export interface WebPublishingOptions {
  /** Use CSS (default true) */
  css?: boolean;
  /** Use thicket format (default true) */
  thicket?: boolean;
  /** Use long file names (default true) */
  longFileNames?: boolean;
  /** Use VML (default false) */
  vml?: boolean;
  /** Allow PNG (default false) */
  allowPng?: boolean;
  /** Target screen size (default "800x600") */
  targetScreenSize?: string;
  /** DPI (default 96) */
  dpi?: number;
  /** Code page */
  codePage?: number;
  /** Character set */
  characterSet?: string;
}

/** File sharing properties (CT_FileSharing) */
export interface FileSharingOptions {
  /** Recommend read-only mode (default false) */
  readOnlyRecommended?: boolean;
  /** User name who has the file locked */
  userName?: string;
  /** Legacy reservation password (hex) */
  reservationPassword?: string;
  /** Modern encryption: algorithm name */
  algorithmName?: string;
  /** Modern encryption: base64 hash value */
  hashValue?: string;
  /** Modern encryption: base64 salt value */
  saltValue?: string;
  /** Modern encryption: spin count */
  spinCount?: number;
}

/** Workbook properties (CT_WorkbookPr) */
export interface WorkbookPrOptions {
  /** Use 1904 date system (default false) */
  date1904?: boolean;
  /** Default theme version */
  defaultThemeVersion?: number;
  /** Show objects: "all" | "placeholders" | "none" */
  showObjects?: string;
  /** Hide pivot field list (default false) */
  hidePivotFieldList?: boolean;
  /** Allow refresh queries (default false) */
  allowRefreshQuery?: boolean;
  /** Filter privacy (default false) */
  filterPrivacy?: boolean;
  /** Backup file (default false) */
  backupFile?: boolean;
  /** Code name */
  codeName?: string;
  /** Show border unselected tables (CT_WorkbookPr @showBorderUnselectedTables) */
  showBorderUnselectedTables?: boolean;
  /** Prompted solutions (CT_WorkbookPr @promptedSolutions) */
  promptedSolutions?: boolean;
  /** Show ink annotation (CT_WorkbookPr @showInkAnnotation) */
  showInkAnnotation?: boolean;
  /** Save external link values (CT_WorkbookPr @saveExternalLinkValues) */
  saveExternalLinkValues?: boolean;
  /** Update links mode (CT_WorkbookPr @updateLinks) */
  updateLinks?: string;
  /** Show pivot chart filter (CT_WorkbookPr @showPivotChartFilter) */
  showPivotChartFilter?: boolean;
  /** Publish items (CT_WorkbookPr @publishItems) */
  publishItems?: boolean;
  /** Check compatibility (CT_WorkbookPr @checkCompatibility) */
  checkCompatibility?: boolean;
  /** Auto compress pictures (CT_WorkbookPr @autoCompressPictures) */
  autoCompressPictures?: boolean;
  /** Refresh all connections (CT_WorkbookPr @refreshAllConnections) */
  refreshAllConnections?: boolean;
}

/** Volatile type entry (CT_VolType) */
export interface VolTypeOptions {
  /** Type of volatile dependency (default: "realTimeData") */
  type?: "realTimeData" | "olapFunctions";
  /** Main volatile dependencies (CT_VolMain, required) */
  mains?: VolMainOptions[];
}

/** Main volatile dependency (CT_VolMain) */
export interface VolMainOptions {
  /** First reference (required) */
  first: string;
  /** Volatile topics (CT_VolTopic) */
  topics?: VolTopicOptions[];
}

/** Volatile topic (CT_VolTopic) */
export interface VolTopicOptions {
  /** Topic value (required) */
  value: string;
  /** Value type (default: "n") */
  valueType?: string;
  /** String topics (stp elements) */
  stringTopics?: string[];
  /** Topic references (CT_VolTopicRef) */
  refs?: VolTopicRefOptions[];
}

/** Volatile topic reference (CT_VolTopicRef) */
export interface VolTopicRefOptions {
  /** Cell reference (required) */
  reference: string;
  /** Sheet index (required) */
  sheetIndex: number;
}

/** Web publish object (CT_WebPublishObject) */
export interface WebPublishObjectOptions {
  /** Relationship ID to the published item */
  rId: string;
  /** Destination file name */
  destinationFile?: string;
  /** Auto republish (default: false) */
  autoRepublish?: boolean;
  /** Title of the published item */
  title?: string;
  /** Source object reference */
  sourceObject?: string;
  /** App name (default: "Excel") */
  appName?: string;
}

/** Calculation properties (CT_CalcPr) */
export interface CalcPrOptions {
  /** Calculation mode: "manual" | "auto" | "autoNoTable" */
  calcMode?: string;
  /** Calc ID (default 162913) */
  calcId?: number;
  /** Full calc on load (default false) */
  fullCalcOnLoad?: boolean;
  /** Calc on save (default true) */
  calcOnSave?: boolean;
  /** Force full calc */
  forceFullCalc?: boolean;
  /** Concurrent calc (default true) */
  concurrentCalc?: boolean;
  /** Concurrent manual count */
  concurrentManualCount?: number;
  /** Iterate (default false) */
  iterate?: boolean;
  /** Iterate count (default 100) */
  iterateCount?: number;
  /** Iterate delta (default 0.001) */
  iterateDelta?: number;
  /** Reference mode: "A1" | "R1C1" */
  refMode?: string;
  /** Full precision (default true) */
  fullPrecision?: boolean;
  /** Calc completed (CT_CalcPr @calcCompleted) */
  calcCompleted?: boolean;
}

/** Workbook view options (CT_BookView) */
export interface WorkbookViewOptions {
  /** Active tab index (0-based) */
  activeTab?: number;
  /** Auto filter date grouping (default true) */
  autoFilterDateGrouping?: boolean;
  /** First sheet tab */
  firstSheet?: number;
  /** Show horizontal scroll (default true) */
  showHorizontalScroll?: boolean;
  /** Show sheet tabs (default true) */
  showSheetTabs?: boolean;
  /** Show vertical scroll (default true) */
  showVerticalScroll?: boolean;
  /** Tab ratio (default 600) */
  tabRatio?: number;
  /** Window width in twips */
  windowWidth?: number;
  /** Window height in twips */
  windowHeight?: number;
  /** X position of the window */
  xWindow?: number;
  /** Y position of the window */
  yWindow?: number;
}

// ── Descriptor Types ──

export interface WorkbookDescriptorOptions {
  sheets: SheetDefinition[];
  pivotCaches?: PivotCacheReference[];
  protection?: WorkbookProtectionOptions;
  customViews?: CustomWorkbookViewOptions[];
  fileRecoveryPr?: FileRecoveryPrOptions;
  functionGroups?: string[];
  webPublishing?: WebPublishingOptions;
  fileSharing?: FileSharingOptions;
  workbookPr?: WorkbookPrOptions;
  calcPr?: CalcPrOptions;
  bookView?: WorkbookViewOptions;
  volTypes?: VolTypeOptions[];
  webPublishObjects?: WebPublishObjectOptions[];
  conformance?: WorkbookConformance;
}

// ── Descriptor ──

export const workbookDesc: CustomDescriptor<WorkbookDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return stringifyWorkbook(opts);
  },

  parse(el, _ctx) {
    const result: Record<string, unknown> = {};

    // Sheets
    const sheetsEl = findChild(el, "sheets");
    if (sheetsEl) {
      const sheets: SheetDefinition[] = [];
      for (const s of sheetsEl.elements ?? []) {
        if (s.name !== "sheet") continue;
        const name = attr(s, "name") ?? "";
        const sheetId = attrNum(s, "sheetId") ?? 0;
        const rId = (s.attributes?.["r:id"] as string | undefined) ?? "";
        const state = attr(s, "state") as SheetDefinition["state"];
        sheets.push({ name, sheetId, rId, state });
      }
      result.sheets = sheets;
    }

    // Pivot caches
    const pivotCachesEl = findChild(el, "pivotCaches");
    if (pivotCachesEl) {
      const caches: PivotCacheReference[] = [];
      for (const pc of pivotCachesEl.elements ?? []) {
        if (pc.name !== "pivotCache") continue;
        caches.push({
          cacheId: attrNum(pc, "cacheId") ?? 0,
          rId: (pc.attributes?.["r:id"] as string) ?? "",
        });
      }
      result.pivotCaches = caches;
    }

    // Workbook protection
    const protEl = findChild(el, "workbookProtection");
    if (protEl?.attributes) {
      const prot: Record<string, unknown> = {};
      if (attr(protEl, "lockStructure") === "1") prot.lockStructure = true;
      if (attr(protEl, "lockWindows") === "1") prot.lockWindows = true;
      if (attr(protEl, "lockRevision") === "1") prot.lockRevision = true;
      if (attr(protEl, "workbookPassword"))
        prot.workbookPassword = attr(protEl, "workbookPassword");
      if (attr(protEl, "workbookAlgorithmName"))
        prot.workbookAlgorithmName = attr(protEl, "workbookAlgorithmName");
      if (attr(protEl, "workbookHashValue"))
        prot.workbookHashValue = attr(protEl, "workbookHashValue");
      if (attr(protEl, "workbookSaltValue"))
        prot.workbookSaltValue = attr(protEl, "workbookSaltValue");
      if (attr(protEl, "workbookSpinCount"))
        prot.workbookSpinCount = attrNum(protEl, "workbookSpinCount");
      result.protection = prot;
    }

    // Book views
    const bookViewsEl = findChild(el, "bookViews");
    if (bookViewsEl) {
      const bvEl = findChild(bookViewsEl, "workbookView");
      if (bvEl?.attributes) {
        const bv: Record<string, unknown> = {};
        const xw = attrNum(bvEl, "xWindow");
        if (xw !== undefined) bv.xWindow = xw;
        const yw = attrNum(bvEl, "yWindow");
        if (yw !== undefined) bv.yWindow = yw;
        const ww = attrNum(bvEl, "windowWidth");
        if (ww !== undefined) bv.windowWidth = ww;
        const wh = attrNum(bvEl, "windowHeight");
        if (wh !== undefined) bv.windowHeight = wh;
        const at = attrNum(bvEl, "activeTab");
        if (at !== undefined) bv.activeTab = at;
        if (attr(bvEl, "showHorizontalScroll") === "0") bv.showHorizontalScroll = false;
        if (attr(bvEl, "showVerticalScroll") === "0") bv.showVerticalScroll = false;
        if (attr(bvEl, "showSheetTabs") === "0") bv.showSheetTabs = false;
        result.bookView = bv;
      }
    }

    // Calc properties
    const calcPrEl = findChild(el, "calcPr");
    if (calcPrEl?.attributes) {
      const calc: Record<string, unknown> = {};
      const calcId = attrNum(calcPrEl, "calcId");
      if (calcId !== undefined) calc.calcId = calcId;
      if (attr(calcPrEl, "calcMode")) calc.calcMode = attr(calcPrEl, "calcMode");
      if (attr(calcPrEl, "fullCalcOnLoad") === "1") calc.fullCalcOnLoad = true;
      if (attr(calcPrEl, "concurrentCalc") === "0") calc.concurrentCalc = false;
      if (attr(calcPrEl, "refMode")) calc.refMode = attr(calcPrEl, "refMode");
      if (attr(calcPrEl, "calcOnSave") === "0") calc.calcOnSave = false;
      if (attr(calcPrEl, "forceFullCalc") === "1") calc.forceFullCalc = true;
      const cmc = attrNum(calcPrEl, "concurrentManualCount");
      if (cmc !== undefined) calc.concurrentManualCount = cmc;
      if (attr(calcPrEl, "iterate") === "1") calc.iterate = true;
      const ic = attrNum(calcPrEl, "iterateCount");
      if (ic !== undefined) calc.iterateCount = ic;
      const id = attrNum(calcPrEl, "iterateDelta");
      if (id !== undefined) calc.iterateDelta = id;
      if (attr(calcPrEl, "fullPrecision") === "0") calc.fullPrecision = false;
      if (attr(calcPrEl, "calcCompleted") === "1") calc.calcCompleted = true;
      result.calcPr = calc;
    }

    // Custom workbook views
    const customViewsEl = findChild(el, "customWorkbookViews");
    if (customViewsEl) {
      const views: CustomWorkbookViewOptions[] = [];
      for (const v of customViewsEl.elements ?? []) {
        if (v.name !== "customWorkbookView") continue;
        const view: Record<string, unknown> = {
          name: attr(v, "name") ?? "",
          guid: attr(v, "guid") ?? "",
          windowWidth: attrNum(v, "windowWidth") ?? 0,
          windowHeight: attrNum(v, "windowHeight") ?? 0,
          activeSheetId: attrNum(v, "activeSheetId") ?? 1,
        };
        const xw = attrNum(v, "xWindow");
        if (xw !== undefined) view.xWindow = xw;
        const yw = attrNum(v, "yWindow");
        if (yw !== undefined) view.yWindow = yw;
        if (attr(v, "showFormulaBar") === "0") view.showFormulaBar = false;
        if (attr(v, "showStatusbar") === "0") view.showStatusbar = false;
        if (attr(v, "showHorizontalScroll") === "0") view.showHorizontalScroll = false;
        if (attr(v, "showVerticalScroll") === "0") view.showVerticalScroll = false;
        if (attr(v, "showSheetTabs") === "0") view.showSheetTabs = false;
        const tabRatio = attrNum(v, "tabRatio");
        if (tabRatio !== undefined) view.tabRatio = tabRatio;
        if (attr(v, "includeHiddenRowCol") === "0") view.includeHiddenRowCol = false;
        if (attr(v, "includePrintSettings") === "0") view.includePrintSettings = false;
        if (attr(v, "personalView") === "1") view.personalView = true;
        if (attr(v, "maximized") === "1") view.maximized = true;
        if (attr(v, "minimized") === "1") view.minimized = true;
        if (attr(v, "autoUpdate") === "1") view.autoUpdate = true;
        const mi = attrNum(v, "mergeInterval");
        if (mi !== undefined) view.mergeInterval = mi;
        if (attr(v, "changesSavedWin") === "1") view.changesSavedWin = true;
        if (attr(v, "onlySync") === "1") view.onlySync = true;
        if (attr(v, "showComments")) view.showComments = attr(v, "showComments");
        views.push(view as unknown as CustomWorkbookViewOptions);
      }
      if (views.length > 0) result.customViews = views;
    }

    // File sharing
    const fileSharingEl = findChild(el, "fileSharing");
    if (fileSharingEl?.attributes) {
      const fs: Record<string, unknown> = {};
      if (attr(fileSharingEl, "readOnlyRecommended") === "1") fs.readOnlyRecommended = true;
      if (attr(fileSharingEl, "userName")) fs.userName = attr(fileSharingEl, "userName");
      if (attr(fileSharingEl, "reservationPassword"))
        fs.reservationPassword = attr(fileSharingEl, "reservationPassword");
      if (attr(fileSharingEl, "algorithmName"))
        fs.algorithmName = attr(fileSharingEl, "algorithmName");
      if (attr(fileSharingEl, "hashValue")) fs.hashValue = attr(fileSharingEl, "hashValue");
      if (attr(fileSharingEl, "saltValue")) fs.saltValue = attr(fileSharingEl, "saltValue");
      const sc = attrNum(fileSharingEl, "spinCount");
      if (sc !== undefined) fs.spinCount = sc;
      result.fileSharing = fs;
    }

    // Web publishing
    const webPublishingEl = findChild(el, "webPublishing");
    if (webPublishingEl?.attributes) {
      const wp: Record<string, unknown> = {};
      if (attr(webPublishingEl, "css") === "0") wp.css = false;
      if (attr(webPublishingEl, "thicket") === "0") wp.thicket = false;
      if (attr(webPublishingEl, "longFileNames") === "0") wp.longFileNames = false;
      if (attr(webPublishingEl, "vml") === "1") wp.vml = true;
      if (attr(webPublishingEl, "allowPng") === "1") wp.allowPng = true;
      if (attr(webPublishingEl, "targetScreenSize"))
        wp.targetScreenSize = attr(webPublishingEl, "targetScreenSize");
      if (attrNum(webPublishingEl, "dpi") !== undefined) wp.dpi = attrNum(webPublishingEl, "dpi");
      if (attrNum(webPublishingEl, "codePage") !== undefined)
        wp.codePage = attrNum(webPublishingEl, "codePage");
      if (attr(webPublishingEl, "characterSet"))
        wp.characterSet = attr(webPublishingEl, "characterSet");
      result.webPublishing = wp;
    }

    // File recovery
    const fileRecoveryEl = findChild(el, "fileRecoveryPr");
    if (fileRecoveryEl?.attributes) {
      const frp: Record<string, unknown> = {};
      if (attr(fileRecoveryEl, "autoRecover") === "0") frp.autoRecover = false;
      if (attr(fileRecoveryEl, "crashSave") === "1") frp.crashSave = true;
      if (attr(fileRecoveryEl, "dataExtractLoad") === "1") frp.dataExtractLoad = true;
      if (attr(fileRecoveryEl, "repairLoad") === "1") frp.repairLoad = true;
      result.fileRecoveryPr = frp;
    }

    // Workbook properties
    const wbPrEl = findChild(el, "workbookPr");
    if (wbPrEl?.attributes) {
      const wbPr: Record<string, unknown> = {};
      if (attr(wbPrEl, "date1904") === "1") wbPr.date1904 = true;
      const dtv = attrNum(wbPrEl, "defaultThemeVersion");
      if (dtv !== undefined) wbPr.defaultThemeVersion = dtv;
      if (attr(wbPrEl, "showObjects")) wbPr.showObjects = attr(wbPrEl, "showObjects");
      if (attr(wbPrEl, "hidePivotFieldList") === "1") wbPr.hidePivotFieldList = true;
      if (attr(wbPrEl, "allowRefreshQuery") === "1") wbPr.allowRefreshQuery = true;
      if (attr(wbPrEl, "filterPrivacy") === "1") wbPr.filterPrivacy = true;
      if (attr(wbPrEl, "backupFile") === "1") wbPr.backupFile = true;
      if (attr(wbPrEl, "codeName")) wbPr.codeName = attr(wbPrEl, "codeName");
      if (attr(wbPrEl, "showBorderUnselectedTables") === "1")
        wbPr.showBorderUnselectedTables = true;
      if (attr(wbPrEl, "promptedSolutions") === "1") wbPr.promptedSolutions = true;
      if (attr(wbPrEl, "showInkAnnotation") === "0") wbPr.showInkAnnotation = false;
      if (attr(wbPrEl, "saveExternalLinkValues") === "0") wbPr.saveExternalLinkValues = false;
      if (attr(wbPrEl, "updateLinks")) wbPr.updateLinks = attr(wbPrEl, "updateLinks");
      if (attr(wbPrEl, "showPivotChartFilter") === "1") wbPr.showPivotChartFilter = true;
      if (attr(wbPrEl, "publishItems") === "1") wbPr.publishItems = true;
      if (attr(wbPrEl, "checkCompatibility") === "1") wbPr.checkCompatibility = true;
      if (attr(wbPrEl, "autoCompressPictures") === "0") wbPr.autoCompressPictures = false;
      if (attr(wbPrEl, "refreshAllConnections") === "1") wbPr.refreshAllConnections = true;
      result.workbookPr = wbPr;
    }

    // Function groups
    const fgEl = findChild(el, "functionGroups");
    if (fgEl) {
      const names: string[] = [];
      for (const fg of fgEl.elements ?? []) {
        if (fg.name === "functionGroup" && attr(fg, "name")) {
          names.push(attr(fg, "name")!);
        }
      }
      if (names.length > 0) result.functionGroups = names;
    }

    // Web publish objects
    const wpoEl = findChild(el, "webPublishObjects");
    if (wpoEl) {
      const objs: WebPublishObjectOptions[] = [];
      for (const wo of wpoEl.elements ?? []) {
        if (wo.name !== "webPublishObject") continue;
        const obj: Record<string, unknown> = {};
        const rId = wo.attributes?.["r:id"] as string | undefined;
        if (rId) obj.rId = rId;
        if (attr(wo, "destinationFile")) obj.destinationFile = attr(wo, "destinationFile");
        if (attr(wo, "autoRepublish") === "1") obj.autoRepublish = true;
        if (attr(wo, "title")) obj.title = attr(wo, "title");
        if (attr(wo, "sourceObject")) obj.sourceObject = attr(wo, "sourceObject");
        if (attr(wo, "appName")) obj.appName = attr(wo, "appName");
        objs.push(obj as unknown as WebPublishObjectOptions);
      }
      if (objs.length > 0) result.webPublishObjects = objs;
    }

    // Volatile types (volTypes)
    const vtEl = findChild(el, "volTypes");
    if (vtEl) {
      const volTypes: VolTypeOptions[] = [];
      for (const vt of vtEl.elements ?? []) {
        if (vt.name !== "volType") continue;
        const volType: Record<string, unknown> = {};
        if (attr(vt, "type")) volType.type = attr(vt, "type") as VolTypeOptions["type"];
        const mains: VolMainOptions[] = [];
        for (const m of vt.elements ?? []) {
          if (m.name !== "main") continue;
          const main: Record<string, unknown> = {};
          if (attr(m, "first")) main.first = attr(m, "first")!;
          const topics: VolTopicOptions[] = [];
          for (const tp of m.elements ?? []) {
            if (tp.name !== "tp") continue;
            const topic: Record<string, unknown> = {};
            const vEl = findChild(tp, "v");
            if (vEl) topic.value = vEl.elements?.[0]?.text ?? "";
            if (attr(tp, "t")) topic.valueType = attr(tp, "t");
            const stps: string[] = [];
            const refs: VolTopicRefOptions[] = [];
            for (const inner of tp.elements ?? []) {
              if (inner.name === "stp") stps.push(String(inner.elements?.[0]?.text ?? ""));
              if (inner.name === "tr") {
                const ref: Record<string, unknown> = {};
                if (attr(inner, "r")) ref.reference = attr(inner, "r")!;
                const sIdx = attrNum(inner, "s");
                if (sIdx !== undefined) ref.sheetIndex = sIdx;
                refs.push(ref as unknown as VolTopicRefOptions);
              }
            }
            if (stps.length > 0) topic.stringTopics = stps;
            if (refs.length > 0) topic.refs = refs;
            topics.push(topic as unknown as VolTopicOptions);
          }
          if (topics.length > 0) main.topics = topics;
          mains.push(main as unknown as VolMainOptions);
        }
        if (mains.length > 0) volType.mains = mains;
        volTypes.push(volType as unknown as VolTypeOptions);
      }
      if (volTypes.length > 0) result.volTypes = volTypes;
    }

    // Conformance
    if (el.attributes?.["conformance"]) {
      result.conformance = attr(el, "conformance") as WorkbookConformance;
    }

    return result as unknown as WorkbookDescriptorOptions;
  },
};

// ── Stringify helpers ──

function stringifyWorkbook(opts: WorkbookDescriptorOptions): string {
  const confAttr = opts.conformance ? ` conformance="${opts.conformance}"` : "";
  const parts: string[] = [
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
      ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
      ' mc:Ignorable="x15 xr xr6 xr10 xr2"' +
      ' xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"' +
      ' xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"' +
      ' xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6"' +
      ' xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10"' +
      ` xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"${confAttr}>`,
    '<fileVersion appName="xl" lastEdited="7" lowestEdited="6" rupBuild="29929"/>',
  ];

  // File sharing (after fileVersion, before workbookPr per XSD sequence)
  if (opts.fileSharing) {
    const fs = opts.fileSharing;
    const fsAttrs: string[] = [];
    if (fs.readOnlyRecommended) fsAttrs.push('readOnlyRecommended="1"');
    if (fs.userName) fsAttrs.push(`userName="${escapeXml(fs.userName)}"`);
    if (fs.reservationPassword) {
      fsAttrs.push(`reservationPassword="${escapeXml(fs.reservationPassword)}"`);
      if (fs.hashValue === undefined) {
        const derived = derivePasswordHash(fs.reservationPassword);
        fsAttrs.push(`algorithmName="${escapeXml(derived.algorithmName)}"`);
        fsAttrs.push(`hashValue="${escapeXml(derived.hashValue)}"`);
        fsAttrs.push(`saltValue="${escapeXml(derived.saltValue)}"`);
        fsAttrs.push(`spinCount="${derived.spinCount}"`);
      }
    }
    if (fs.algorithmName) fsAttrs.push(`algorithmName="${escapeXml(fs.algorithmName)}"`);
    if (fs.hashValue) fsAttrs.push(`hashValue="${escapeXml(fs.hashValue)}"`);
    if (fs.saltValue) fsAttrs.push(`saltValue="${escapeXml(fs.saltValue)}"`);
    if (fs.spinCount !== undefined) fsAttrs.push(`spinCount="${fs.spinCount}"`);
    if (fsAttrs.length > 0) {
      parts.push(`<fileSharing ${fsAttrs.join(" ")}/>`);
    }
  }

  // Workbook properties
  if (opts.workbookPr) {
    const wbPr = opts.workbookPr;
    const wbPrAttrs: string[] = [];
    if (wbPr.date1904) wbPrAttrs.push('date1904="1"');
    if (wbPr.defaultThemeVersion !== undefined)
      wbPrAttrs.push(`defaultThemeVersion="${wbPr.defaultThemeVersion}"`);
    if (wbPr.showObjects) wbPrAttrs.push(`showObjects="${escapeXml(wbPr.showObjects)}"`);
    if (wbPr.hidePivotFieldList) wbPrAttrs.push('hidePivotFieldList="1"');
    if (wbPr.allowRefreshQuery) wbPrAttrs.push('allowRefreshQuery="1"');
    if (wbPr.filterPrivacy) wbPrAttrs.push('filterPrivacy="1"');
    if (wbPr.backupFile) wbPrAttrs.push('backupFile="1"');
    if (wbPr.codeName) wbPrAttrs.push(`codeName="${escapeXml(wbPr.codeName)}"`);
    if (wbPr.showBorderUnselectedTables) wbPrAttrs.push('showBorderUnselectedTables="1"');
    if (wbPr.promptedSolutions) wbPrAttrs.push('promptedSolutions="1"');
    if (wbPr.showInkAnnotation === false) wbPrAttrs.push('showInkAnnotation="0"');
    if (wbPr.saveExternalLinkValues === false) wbPrAttrs.push('saveExternalLinkValues="0"');
    if (wbPr.updateLinks) wbPrAttrs.push(`updateLinks="${escapeXml(wbPr.updateLinks)}"`);
    if (wbPr.showPivotChartFilter) wbPrAttrs.push('showPivotChartFilter="1"');
    if (wbPr.publishItems) wbPrAttrs.push('publishItems="1"');
    if (wbPr.checkCompatibility) wbPrAttrs.push('checkCompatibility="1"');
    if (wbPr.autoCompressPictures === false) wbPrAttrs.push('autoCompressPictures="0"');
    if (wbPr.refreshAllConnections) wbPrAttrs.push('refreshAllConnections="1"');
    parts.push(`<workbookPr${wbPrAttrs.length > 0 ? ` ${wbPrAttrs.join(" ")}` : ""}/>`);
  } else {
    parts.push("<workbookPr/>");
  }

  // Workbook protection (after workbookPr, before bookViews per XSD sequence)
  if (opts.protection) {
    const prot = opts.protection;
    const protAttrs: string[] = [];
    if (prot.lockStructure) protAttrs.push('lockStructure="1"');
    if (prot.lockWindows) protAttrs.push('lockWindows="1"');
    if (prot.lockRevision) protAttrs.push('lockRevision="1"');
    if (prot.workbookPassword) {
      protAttrs.push(`workbookPassword="${hashPassword(prot.workbookPassword)}"`);
      if (prot.workbookHashValue === undefined) {
        const wbDerived = derivePasswordHash(prot.workbookPassword);
        protAttrs.push(`workbookAlgorithmName="${escapeXml(wbDerived.algorithmName)}"`);
        protAttrs.push(`workbookHashValue="${escapeXml(wbDerived.hashValue)}"`);
        protAttrs.push(`workbookSaltValue="${escapeXml(wbDerived.saltValue)}"`);
        protAttrs.push(`workbookSpinCount="${wbDerived.spinCount}"`);
      }
    }
    if (prot.workbookAlgorithmName)
      protAttrs.push(`workbookAlgorithmName="${escapeXml(prot.workbookAlgorithmName)}"`);
    if (prot.workbookHashValue)
      protAttrs.push(`workbookHashValue="${escapeXml(prot.workbookHashValue)}"`);
    if (prot.workbookSaltValue)
      protAttrs.push(`workbookSaltValue="${escapeXml(prot.workbookSaltValue)}"`);
    if (prot.workbookSpinCount !== undefined)
      protAttrs.push(`workbookSpinCount="${prot.workbookSpinCount}"`);
    if (prot.revisionsPassword) {
      protAttrs.push(`revisionsPassword="${hashPassword(prot.revisionsPassword)}"`);
      if (prot.revisionsHashValue === undefined) {
        const revDerived = derivePasswordHash(prot.revisionsPassword);
        protAttrs.push(`revisionsAlgorithmName="${escapeXml(revDerived.algorithmName)}"`);
        protAttrs.push(`revisionsHashValue="${escapeXml(revDerived.hashValue)}"`);
        protAttrs.push(`revisionsSaltValue="${escapeXml(revDerived.saltValue)}"`);
        protAttrs.push(`revisionsSpinCount="${revDerived.spinCount}"`);
      }
    }
    if (prot.revisionsAlgorithmName)
      protAttrs.push(`revisionsAlgorithmName="${escapeXml(prot.revisionsAlgorithmName)}"`);
    if (prot.revisionsHashValue)
      protAttrs.push(`revisionsHashValue="${escapeXml(prot.revisionsHashValue)}"`);
    if (prot.revisionsSaltValue)
      protAttrs.push(`revisionsSaltValue="${escapeXml(prot.revisionsSaltValue)}"`);
    if (prot.revisionsSpinCount !== undefined)
      protAttrs.push(`revisionsSpinCount="${prot.revisionsSpinCount}"`);
    if (prot.workbookPasswordCharacterSet)
      protAttrs.push(
        `workbookPasswordCharacterSet="${escapeXml(prot.workbookPasswordCharacterSet)}"`,
      );
    if (prot.revisionsPasswordCharacterSet)
      protAttrs.push(
        `revisionsPasswordCharacterSet="${escapeXml(prot.revisionsPasswordCharacterSet)}"`,
      );
    if (protAttrs.length > 0) {
      parts.push(`<workbookProtection ${protAttrs.join(" ")}/>`);
    }
  }

  // Book views
  if (opts.bookView) {
    const bv = opts.bookView;
    const bvAttrs: string[] = [];
    if (bv.xWindow !== undefined) bvAttrs.push(`xWindow="${bv.xWindow}"`);
    else bvAttrs.push('xWindow="0"');
    if (bv.yWindow !== undefined) bvAttrs.push(`yWindow="${bv.yWindow}"`);
    else bvAttrs.push('yWindow="0"');
    if (bv.windowWidth !== undefined) bvAttrs.push(`windowWidth="${bv.windowWidth}"`);
    else bvAttrs.push('windowWidth="28800"');
    if (bv.windowHeight !== undefined) bvAttrs.push(`windowHeight="${bv.windowHeight}"`);
    else bvAttrs.push('windowHeight="12300"');
    if (bv.activeTab !== undefined) bvAttrs.push(`activeTab="${bv.activeTab}"`);
    if (bv.autoFilterDateGrouping === false) bvAttrs.push('autoFilterDateGrouping="0"');
    if (bv.firstSheet !== undefined) bvAttrs.push(`firstSheet="${bv.firstSheet}"`);
    if (bv.showHorizontalScroll === false) bvAttrs.push('showHorizontalScroll="0"');
    if (bv.showSheetTabs === false) bvAttrs.push('showSheetTabs="0"');
    if (bv.showVerticalScroll === false) bvAttrs.push('showVerticalScroll="0"');
    if (bv.tabRatio !== undefined) bvAttrs.push(`tabRatio="${bv.tabRatio}"`);
    parts.push(`<bookViews><workbookView ${bvAttrs.join(" ")}/></bookViews>`);
  } else {
    parts.push(
      '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="12300"/></bookViews>',
    );
  }

  parts.push("<sheets>");
  for (const s of opts.sheets) {
    const stateAttr = s.state && s.state !== "visible" ? ` state="${s.state}"` : "";
    parts.push(
      `<sheet name="${escapeXml(s.name)}" sheetId="${s.sheetId}" r:id="${s.rId}"${stateAttr}/>`,
    );
  }
  parts.push("</sheets>");

  // Function groups (after sheets, before externalReferences per XSD)
  const functionGroups = opts.functionGroups ?? [];
  if (functionGroups.length > 0) {
    const fgParts: string[] = [`<functionGroups builtInGroupCount="16">`];
    for (const name of functionGroups) {
      fgParts.push(`<functionGroup name="${escapeXml(name)}"/>`);
    }
    fgParts.push("</functionGroups>");
    parts.push(fgParts.join(""));
  }

  // externalReferences placeholder — compiler injects the XML here if needed
  parts.push("<!--EXTERNAL_REFS-->");

  // Calculation properties
  if (opts.calcPr) {
    const cp = opts.calcPr;
    const cpAttrs: string[] = [];
    cpAttrs.push(`calcId="${cp.calcId ?? 162913}"`);
    if (cp.calcMode) cpAttrs.push(`calcMode="${escapeXml(cp.calcMode)}"`);
    if (cp.fullCalcOnLoad) cpAttrs.push('fullCalcOnLoad="1"');
    if (cp.calcOnSave === false) cpAttrs.push('calcOnSave="0"');
    if (cp.forceFullCalc) cpAttrs.push('forceFullCalc="1"');
    if (cp.concurrentCalc === false) cpAttrs.push('concurrentCalc="0"');
    if (cp.concurrentManualCount !== undefined)
      cpAttrs.push(`concurrentManualCount="${cp.concurrentManualCount}"`);
    if (cp.iterate) cpAttrs.push('iterate="1"');
    if (cp.iterateCount !== undefined) cpAttrs.push(`iterateCount="${cp.iterateCount}"`);
    if (cp.iterateDelta !== undefined) cpAttrs.push(`iterateDelta="${cp.iterateDelta}"`);
    if (cp.refMode) cpAttrs.push(`refMode="${escapeXml(cp.refMode)}"`);
    if (cp.fullPrecision === false) cpAttrs.push('fullPrecision="0"');
    if (cp.calcCompleted) cpAttrs.push('calcCompleted="1"');
    parts.push(`<calcPr ${cpAttrs.join(" ")}/>`);
  } else {
    parts.push('<calcPr calcId="162913"/>');
  }

  // Custom workbook views (after calcPr, before pivotCaches per XSD)
  if (opts.customViews && opts.customViews.length > 0) {
    parts.push("<customWorkbookViews>");
    for (const v of opts.customViews) {
      const vAttrs: string[] = [
        `name="${escapeXml(v.name)}"`,
        `guid="${escapeXml(v.guid)}"`,
        `windowWidth="${v.windowWidth}"`,
        `windowHeight="${v.windowHeight}"`,
        `activeSheetId="${v.activeSheetId}"`,
      ];
      if (v.xWindow !== undefined) vAttrs.push(`xWindow="${v.xWindow}"`);
      if (v.yWindow !== undefined) vAttrs.push(`yWindow="${v.yWindow}"`);
      if (v.showFormulaBar === false) vAttrs.push('showFormulaBar="0"');
      if (v.showStatusbar === false) vAttrs.push('showStatusbar="0"');
      if (v.showHorizontalScroll === false) vAttrs.push('showHorizontalScroll="0"');
      if (v.showVerticalScroll === false) vAttrs.push('showVerticalScroll="0"');
      if (v.showSheetTabs === false) vAttrs.push('showSheetTabs="0"');
      if (v.tabRatio !== undefined) vAttrs.push(`tabRatio="${v.tabRatio}"`);
      if (v.includeHiddenRowCol === false) vAttrs.push('includeHiddenRowCol="0"');
      if (v.includePrintSettings === false) vAttrs.push('includePrintSettings="0"');
      if (v.personalView) vAttrs.push('personalView="1"');
      if (v.maximized) vAttrs.push('maximized="1"');
      if (v.minimized) vAttrs.push('minimized="1"');
      if (v.autoUpdate) vAttrs.push('autoUpdate="1"');
      if (v.mergeInterval !== undefined) vAttrs.push(`mergeInterval="${v.mergeInterval}"`);
      if (v.changesSavedWin) vAttrs.push('changesSavedWin="1"');
      if (v.onlySync) vAttrs.push('onlySync="1"');
      if (v.showComments) vAttrs.push(`showComments="${escapeXml(v.showComments)}"`);
      parts.push(`<customWorkbookView ${vAttrs.join(" ")}/>`);
    }
    parts.push("</customWorkbookViews>");
  }

  const pivotCaches = opts.pivotCaches ?? [];
  if (pivotCaches.length > 0) {
    parts.push("<pivotCaches>");
    for (const pc of pivotCaches) {
      parts.push(`<pivotCache cacheId="${pc.cacheId}" r:id="${pc.rId}"/>`);
    }
    parts.push("</pivotCaches>");
  }

  // Web publishing (after pivotCaches, before fileRecoveryPr per XSD sequence)
  if (opts.webPublishing) {
    const wp = opts.webPublishing;
    const wpAttrs: string[] = [];
    if (wp.css === false) wpAttrs.push('css="0"');
    if (wp.thicket === false) wpAttrs.push('thicket="0"');
    if (wp.longFileNames === false) wpAttrs.push('longFileNames="0"');
    if (wp.vml) wpAttrs.push('vml="1"');
    if (wp.allowPng) wpAttrs.push('allowPng="1"');
    if (wp.targetScreenSize && wp.targetScreenSize !== "800x600")
      wpAttrs.push(`targetScreenSize="${wp.targetScreenSize}"`);
    if (wp.dpi !== undefined && wp.dpi !== 96) wpAttrs.push(`dpi="${wp.dpi}"`);
    if (wp.codePage !== undefined) wpAttrs.push(`codePage="${wp.codePage}"`);
    if (wp.characterSet) wpAttrs.push(`characterSet="${escapeXml(wp.characterSet)}"`);
    parts.push(`<webPublishing ${wpAttrs.join(" ")}/>`);
  }

  // File recovery properties (after webPublishing per XSD sequence)
  if (opts.fileRecoveryPr) {
    const frp = opts.fileRecoveryPr;
    const frpAttrs: string[] = [];
    if (frp.autoRecover === false) frpAttrs.push('autoRecover="0"');
    if (frp.crashSave) frpAttrs.push('crashSave="1"');
    if (frp.dataExtractLoad) frpAttrs.push('dataExtractLoad="1"');
    if (frp.repairLoad) frpAttrs.push('repairLoad="1"');
    if (frpAttrs.length > 0) {
      parts.push(`<fileRecoveryPr ${frpAttrs.join(" ")}/>`);
    }
  }

  // Web publish objects (after fileRecoveryPr per XSD sequence)
  if (opts.webPublishObjects && opts.webPublishObjects.length > 0) {
    const wpoParts: string[] = [`<webPublishObjects count="${opts.webPublishObjects.length}">`];
    for (const wpo of opts.webPublishObjects) {
      const wpoAttrs: string[] = [`r:id="${escapeXml(wpo.rId)}"`];
      if (wpo.destinationFile) wpoAttrs.push(`destinationFile="${escapeXml(wpo.destinationFile)}"`);
      if (wpo.autoRepublish) wpoAttrs.push('autoRepublish="1"');
      if (wpo.title) wpoAttrs.push(`title="${escapeXml(wpo.title)}"`);
      if (wpo.sourceObject) wpoAttrs.push(`sourceObject="${escapeXml(wpo.sourceObject)}"`);
      wpoParts.push(`<webPublishObject ${wpoAttrs.join(" ")}/>`);
    }
    wpoParts.push("</webPublishObjects>");
    parts.push(wpoParts.join(""));
  }

  // Volatile dependencies (volTypes)
  if (opts.volTypes && opts.volTypes.length > 0) {
    const vtParts: string[] = [`<volTypes count="${opts.volTypes.length}">`];
    for (const vt of opts.volTypes) {
      const vtType = vt.type ?? "realTimeData";
      const mains = vt.mains ?? [];
      if (mains.length > 0) {
        const mainParts: string[] = [];
        for (const m of mains) {
          const tpParts: string[] = [];
          for (const topic of m.topics ?? []) {
            const tpInner: string[] = [`<v>${escapeXml(topic.value)}</v>`];
            for (const stp of topic.stringTopics ?? []) {
              tpInner.push(`<stp>${escapeXml(stp)}</stp>`);
            }
            for (const tr of topic.refs ?? []) {
              tpInner.push(`<tr r="${escapeXml(tr.reference)}" s="${tr.sheetIndex}"/>`);
            }
            const tpAttr =
              topic.valueType && topic.valueType !== "n"
                ? ` t="${escapeXml(topic.valueType)}"`
                : "";
            tpParts.push(`<tp${tpAttr}>${tpInner.join("")}</tp>`);
          }
          mainParts.push(`<main first="${escapeXml(m.first)}">${tpParts.join("")}</main>`);
        }
        vtParts.push(`<volType type="${vtType}">${mainParts.join("")}</volType>`);
      } else {
        vtParts.push(`<volType type="${vtType}"/>`);
      }
    }
    vtParts.push("</volTypes>");
    parts.push(vtParts.join(""));
  }

  parts.push("</workbook>");
  return parts.join("");
}

// ── Exported helper functions ──

/** Generate tableParts XML fragment for embedding in a worksheet. */
export function buildTablePartsXml(tableParts: TablePartReference[]): string {
  if (tableParts.length === 0) return "";
  const p: string[] = [`<tableParts count="${tableParts.length}">`];
  for (const tp of tableParts) {
    p.push(`<tablePart r:id="${tp.rId}"/>`);
  }
  p.push("</tableParts>");
  return p.join("");
}

/** Generate externalReferences XML fragment for embedding in the workbook. */
export function buildExternalReferencesXml(refs: { rId: string }[]): string {
  if (refs.length === 0) return "";
  const p: string[] = ["<externalReferences>"];
  for (const ref of refs) {
    p.push(`<externalReference r:id="${ref.rId}"/>`);
  }
  p.push("</externalReferences>");
  return p.join("");
}

/** Legacy Excel password hash (XOR-based) */
function hashPassword(password: string): string {
  let hash = 0;
  for (let i = 0; i < password.length; i++) {
    const c = password.charCodeAt(i);
    hash = ((hash >> 14) & 1) + ((hash << 1) & 0x7fff);
    hash ^= c;
    hash = hash & 0x4000 ? hash ^ 0x1 : hash;
  }
  hash = ((hash >> 14) & 1) + ((hash << 1) & 0x7fff);
  hash = ((hash >> 14) & 1) + ((hash << 1) & 0x7fff);
  hash ^= password.length;
  return hash.toString(16).toUpperCase().padStart(4, "0");
}
