/**
 * Workbook types and descriptor for SpreadsheetML documents.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrNum, findChild } from "@office-open/xml";

import { stringifyWorkbook } from "./stringify";
import type {
  CalculationPropertiesOptions,
  CustomWorkbookViewOptions,
  DefinedNameOptions,
  FileRecoveryPropertiesOptions,
  FileSharingOptions,
  FileVersionOptions,
  PivotCacheReference,
  SheetDefinition,
  SmartTagPropertiesOptions,
  SmartTagShow,
  SmartTagTypeOptions,
  VolMainOptions,
  VolTopicOptions,
  VolTopicRefOptions,
  VolTypeOptions,
  WebPublishObjectOptions,
  WebPublishingOptions,
  WorkbookConformance,
  WorkbookDescriptorOptions,
  WorkbookProtectionOptions,
  WorkbookPropertiesOptions,
  WorkbookViewOptions,
} from "./types";

// ── Descriptor ──

export const workbookDesc: CustomDescriptor<WorkbookDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return stringifyWorkbook(opts);
  },

  parse(el, _ctx) {
    const result: Partial<WorkbookDescriptorOptions> = {};

    // File version (CT_Workbook first child) — Excel version stamp
    const fileVersionEl = findChild(el, "fileVersion");
    if (fileVersionEl) {
      const fv: FileVersionOptions = {};
      const appName = attr(fileVersionEl, "appName");
      if (appName) fv.appName = appName;
      const lastEdited = attrNum(fileVersionEl, "lastEdited");
      if (lastEdited !== undefined) fv.lastEdited = lastEdited;
      const lowestEdited = attrNum(fileVersionEl, "lowestEdited");
      if (lowestEdited !== undefined) fv.lowestEdited = lowestEdited;
      const rupBuild = attrNum(fileVersionEl, "rupBuild");
      if (rupBuild !== undefined) fv.rupBuild = rupBuild;
      result.fileVersion = fv;
    }

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
      const prot: WorkbookProtectionOptions = {};
      if (String(attr(protEl, "lockStructure")) === "1") prot.lockStructure = true;
      if (String(attr(protEl, "lockWindows")) === "1") prot.lockWindows = true;
      if (String(attr(protEl, "lockRevision")) === "1") prot.lockRevision = true;
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
        const bv: WorkbookViewOptions = {};
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
        if (String(attr(bvEl, "autoFilterDateGrouping")) === "0") bv.autoFilterDateGrouping = false;
        const fs = attrNum(bvEl, "firstSheet");
        if (fs !== undefined) bv.firstSheet = fs;
        if (String(attr(bvEl, "showHorizontalScroll")) === "0") bv.showHorizontalScroll = false;
        if (String(attr(bvEl, "showVerticalScroll")) === "0") bv.showVerticalScroll = false;
        if (String(attr(bvEl, "showSheetTabs")) === "0") bv.showSheetTabs = false;
        const tr = attrNum(bvEl, "tabRatio");
        if (tr !== undefined) bv.tabRatio = tr;
        result.bookView = bv;
      }
    }

    // Calc properties
    const calcPrEl = findChild(el, "calcPr");
    if (calcPrEl?.attributes) {
      const calc: CalculationPropertiesOptions = {};
      const calcId = attrNum(calcPrEl, "calcId");
      if (calcId !== undefined) calc.calcId = calcId;
      if (attr(calcPrEl, "calcMode")) calc.calcMode = attr(calcPrEl, "calcMode");
      if (String(attr(calcPrEl, "fullCalcOnLoad")) === "1") calc.fullCalcOnLoad = true;
      if (String(attr(calcPrEl, "concurrentCalc")) === "0") calc.concurrentCalc = false;
      if (attr(calcPrEl, "refMode")) calc.refMode = attr(calcPrEl, "refMode");
      if (String(attr(calcPrEl, "calcOnSave")) === "0") calc.calcOnSave = false;
      if (String(attr(calcPrEl, "forceFullCalc")) === "1") calc.forceFullCalc = true;
      const cmc = attrNum(calcPrEl, "concurrentManualCount");
      if (cmc !== undefined) calc.concurrentManualCount = cmc;
      if (String(attr(calcPrEl, "iterate")) === "1") calc.iterate = true;
      const ic = attrNum(calcPrEl, "iterateCount");
      if (ic !== undefined) calc.iterateCount = ic;
      const id = attrNum(calcPrEl, "iterateDelta");
      if (id !== undefined) calc.iterateDelta = id;
      if (String(attr(calcPrEl, "fullPrecision")) === "0") calc.fullPrecision = false;
      if (String(attr(calcPrEl, "calcCompleted")) === "1") calc.calcCompleted = true;
      result.calcPr = calc;
    }

    // Custom workbook views
    const customViewsEl = findChild(el, "customWorkbookViews");
    if (customViewsEl) {
      const views: CustomWorkbookViewOptions[] = [];
      for (const v of customViewsEl.elements ?? []) {
        if (v.name !== "customWorkbookView") continue;
        const view: CustomWorkbookViewOptions = {
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
        if (String(attr(v, "showFormulaBar")) === "0") view.showFormulaBar = false;
        if (String(attr(v, "showStatusbar")) === "0") view.showStatusbar = false;
        if (String(attr(v, "showHorizontalScroll")) === "0") view.showHorizontalScroll = false;
        if (String(attr(v, "showVerticalScroll")) === "0") view.showVerticalScroll = false;
        if (String(attr(v, "showSheetTabs")) === "0") view.showSheetTabs = false;
        const tabRatio = attrNum(v, "tabRatio");
        if (tabRatio !== undefined) view.tabRatio = tabRatio;
        if (String(attr(v, "includeHiddenRowCol")) === "0") view.includeHiddenRowCol = false;
        if (String(attr(v, "includePrintSettings")) === "0") view.includePrintSettings = false;
        if (String(attr(v, "personalView")) === "1") view.personalView = true;
        if (String(attr(v, "maximized")) === "1") view.maximized = true;
        if (String(attr(v, "minimized")) === "1") view.minimized = true;
        if (String(attr(v, "autoUpdate")) === "1") view.autoUpdate = true;
        const mi = attrNum(v, "mergeInterval");
        if (mi !== undefined) view.mergeInterval = mi;
        if (String(attr(v, "changesSavedWin")) === "1") view.changesSavedWin = true;
        if (String(attr(v, "onlySync")) === "1") view.onlySync = true;
        if (attr(v, "showComments")) view.showComments = attr(v, "showComments");
        views.push(view);
      }
      if (views.length > 0) result.customViews = views;
    }

    // File sharing
    const fileSharingEl = findChild(el, "fileSharing");
    if (fileSharingEl?.attributes) {
      const fs: FileSharingOptions = {};
      if (String(attr(fileSharingEl, "readOnlyRecommended")) === "1") fs.readOnlyRecommended = true;
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
      const wp: WebPublishingOptions = {};
      if (String(attr(webPublishingEl, "css")) === "0") wp.css = false;
      if (String(attr(webPublishingEl, "thicket")) === "0") wp.thicket = false;
      if (String(attr(webPublishingEl, "longFileNames")) === "0") wp.longFileNames = false;
      if (String(attr(webPublishingEl, "vml")) === "1") wp.vml = true;
      if (String(attr(webPublishingEl, "allowPng")) === "1") wp.allowPng = true;
      if (attr(webPublishingEl, "targetScreenSize"))
        wp.targetScreenSize = attr(webPublishingEl, "targetScreenSize");
      const dpi = attrNum(webPublishingEl, "dpi");
      if (dpi !== undefined) wp.dpi = dpi;
      const codePage = attrNum(webPublishingEl, "codePage");
      if (codePage !== undefined) wp.codePage = codePage;
      if (attr(webPublishingEl, "characterSet"))
        wp.characterSet = attr(webPublishingEl, "characterSet");
      result.webPublishing = wp;
    }

    // File recovery
    const fileRecoveryEl = findChild(el, "fileRecoveryPr");
    if (fileRecoveryEl?.attributes) {
      const frp: FileRecoveryPropertiesOptions = {};
      if (String(attr(fileRecoveryEl, "autoRecover")) === "0") frp.autoRecover = false;
      if (String(attr(fileRecoveryEl, "crashSave")) === "1") frp.crashSave = true;
      if (String(attr(fileRecoveryEl, "dataExtractLoad")) === "1") frp.dataExtractLoad = true;
      if (String(attr(fileRecoveryEl, "repairLoad")) === "1") frp.repairLoad = true;
      result.fileRecoveryPr = frp;
    }

    // Workbook properties
    const wbPrEl = findChild(el, "workbookPr");
    if (wbPrEl?.attributes) {
      const wbPr: WorkbookPropertiesOptions = {};
      if (String(attr(wbPrEl, "date1904")) === "1") wbPr.date1904 = true;
      const dtv = attrNum(wbPrEl, "defaultThemeVersion");
      if (dtv !== undefined) wbPr.defaultThemeVersion = dtv;
      if (attr(wbPrEl, "showObjects")) wbPr.showObjects = attr(wbPrEl, "showObjects");
      if (String(attr(wbPrEl, "hidePivotFieldList")) === "1") wbPr.hidePivotFieldList = true;
      if (String(attr(wbPrEl, "allowRefreshQuery")) === "1") wbPr.allowRefreshQuery = true;
      if (String(attr(wbPrEl, "filterPrivacy")) === "1") wbPr.filterPrivacy = true;
      if (String(attr(wbPrEl, "backupFile")) === "1") wbPr.backupFile = true;
      if (attr(wbPrEl, "codeName")) wbPr.codeName = attr(wbPrEl, "codeName");
      if (String(attr(wbPrEl, "showBorderUnselectedTables")) === "1")
        wbPr.showBorderUnselectedTables = true;
      if (String(attr(wbPrEl, "promptedSolutions")) === "1") wbPr.promptedSolutions = true;
      if (String(attr(wbPrEl, "showInkAnnotation")) === "0") wbPr.showInkAnnotation = false;
      if (String(attr(wbPrEl, "saveExternalLinkValues")) === "0")
        wbPr.saveExternalLinkValues = false;
      if (attr(wbPrEl, "updateLinks")) wbPr.updateLinks = attr(wbPrEl, "updateLinks");
      if (String(attr(wbPrEl, "showPivotChartFilter")) === "1") wbPr.showPivotChartFilter = true;
      if (String(attr(wbPrEl, "publishItems")) === "1") wbPr.publishItems = true;
      if (String(attr(wbPrEl, "checkCompatibility")) === "1") wbPr.checkCompatibility = true;
      if (String(attr(wbPrEl, "autoCompressPictures")) === "0") wbPr.autoCompressPictures = false;
      if (String(attr(wbPrEl, "refreshAllConnections")) === "1") wbPr.refreshAllConnections = true;
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
        const obj: WebPublishObjectOptions = {
          rId: (wo.attributes?.["r:id"] as string) ?? "",
        };
        if (attr(wo, "destinationFile")) obj.destinationFile = attr(wo, "destinationFile");
        if (String(attr(wo, "autoRepublish")) === "1") obj.autoRepublish = true;
        if (attr(wo, "title")) obj.title = attr(wo, "title");
        if (attr(wo, "sourceObject")) obj.sourceObject = attr(wo, "sourceObject");
        if (attr(wo, "appName")) obj.appName = attr(wo, "appName");
        objs.push(obj);
      }
      if (objs.length > 0) result.webPublishObjects = objs;
    }

    // Volatile types (volTypes)
    const vtEl = findChild(el, "volTypes");
    if (vtEl) {
      const volTypes: VolTypeOptions[] = [];
      for (const vt of vtEl.elements ?? []) {
        if (vt.name !== "volType") continue;
        const volType: VolTypeOptions = {};
        const typeVal = attr(vt, "type");
        if (typeVal) volType.type = typeVal as VolTypeOptions["type"];
        const mains: VolMainOptions[] = [];
        for (const m of vt.elements ?? []) {
          if (m.name !== "main") continue;
          const main: VolMainOptions = { first: attr(m, "first") ?? "" };
          const topics: VolTopicOptions[] = [];
          for (const tp of m.elements ?? []) {
            if (tp.name !== "tp") continue;
            const vEl = findChild(tp, "v");
            const topic: VolTopicOptions = { value: String(vEl?.elements?.[0]?.text ?? "") };
            const tVal = attr(tp, "t");
            if (tVal) topic.valueType = tVal;
            const stps: string[] = [];
            const refs: VolTopicRefOptions[] = [];
            for (const inner of tp.elements ?? []) {
              if (inner.name === "stp") stps.push(String(inner.elements?.[0]?.text ?? ""));
              if (inner.name === "tr") {
                const ref: VolTopicRefOptions = {
                  reference: attr(inner, "r") ?? "",
                  sheetIndex: attrNum(inner, "s") ?? 0,
                };
                refs.push(ref);
              }
            }
            if (stps.length > 0) topic.stringTopics = stps;
            if (refs.length > 0) topic.refs = refs;
            topics.push(topic);
          }
          if (topics.length > 0) main.topics = topics;
          mains.push(main);
        }
        if (mains.length > 0) volType.mains = mains;
        volTypes.push(volType);
      }
      if (volTypes.length > 0) result.volTypes = volTypes;
    }

    // Defined names (CT_DefinedNames — simpleContent ST_Formula + 14 attrs)
    const definedNamesEl = findChild(el, "definedNames");
    if (definedNamesEl) {
      const names: DefinedNameOptions[] = [];
      for (const d of definedNamesEl.elements ?? []) {
        if (d.name !== "definedName") continue;
        const value = (d.elements ?? []).map((e) => e.text ?? "").join("");
        const dn: DefinedNameOptions = { name: attr(d, "name") ?? "", value };
        if (attr(d, "comment") !== undefined) dn.comment = attr(d, "comment");
        if (attr(d, "customMenu") !== undefined) dn.customMenu = attr(d, "customMenu");
        if (attr(d, "description") !== undefined) dn.description = attr(d, "description");
        if (attr(d, "help") !== undefined) dn.help = attr(d, "help");
        if (attr(d, "statusBar") !== undefined) dn.statusBar = attr(d, "statusBar");
        const localSheetId = attrNum(d, "localSheetId");
        if (localSheetId !== undefined) dn.localSheetId = localSheetId;
        if (String(attr(d, "hidden")) === "1") dn.hidden = true;
        if (String(attr(d, "function")) === "1") dn.function = true;
        if (String(attr(d, "vbProcedure")) === "1") dn.vbProcedure = true;
        if (String(attr(d, "xlm")) === "1") dn.xlm = true;
        const functionGroupId = attrNum(d, "functionGroupId");
        if (functionGroupId !== undefined) dn.functionGroupId = functionGroupId;
        if (attr(d, "shortcutKey") !== undefined) dn.shortcutKey = attr(d, "shortcutKey");
        if (String(attr(d, "publishToServer")) === "1") dn.publishToServer = true;
        if (String(attr(d, "workbookParameter")) === "1") dn.workbookParameter = true;
        names.push(dn);
      }
      if (names.length > 0) result.definedNames = names;
    }

    // Smart tag properties (CT_SmartTagPr)
    const smartTagPrEl = findChild(el, "smartTagPr");
    if (smartTagPrEl?.attributes) {
      const stp: SmartTagPropertiesOptions = {};
      if (String(attr(smartTagPrEl, "embed")) === "1") stp.embed = true;
      const show = attr(smartTagPrEl, "show");
      if (show && show !== "all") stp.show = show as SmartTagShow;
      result.smartTagPr = stp;
    }

    // Smart tag types (CT_SmartTagTypes → smartTagType[])
    const smartTagTypesEl = findChild(el, "smartTagTypes");
    if (smartTagTypesEl) {
      const types: SmartTagTypeOptions[] = [];
      for (const t of smartTagTypesEl.elements ?? []) {
        if (t.name !== "smartTagType") continue;
        const type: SmartTagTypeOptions = {};
        const ns = attr(t, "namespaceUri");
        if (ns) type.namespaceUri = ns;
        const n = attr(t, "name");
        if (n) type.name = n;
        const u = attr(t, "url");
        if (u) type.url = u;
        types.push(type);
      }
      if (types.length > 0) result.smartTagTypes = types;
    }

    // Conformance
    if (el.attributes?.["conformance"]) {
      result.conformance = attr(el, "conformance") as WorkbookConformance;
    }

    return result as WorkbookDescriptorOptions;
  },
};
