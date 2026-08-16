/**
 * Workbook types and descriptor for SpreadsheetML documents.
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
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
      if (parseOnOff(attr(protEl, "lockStructure"))) prot.lockStructure = true;
      if (parseOnOff(attr(protEl, "lockWindows"))) prot.lockWindows = true;
      if (parseOnOff(attr(protEl, "lockRevision"))) prot.lockRevision = true;
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
        const vis = attr(bvEl, "visibility");
        if (vis !== undefined) bv.visibility = vis as WorkbookViewOptions["visibility"];
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
      if (parseOnOff(attr(calcPrEl, "fullCalcOnLoad"))) calc.fullCalcOnLoad = true;
      if (String(attr(calcPrEl, "concurrentCalc")) === "0") calc.concurrentCalc = false;
      if (attr(calcPrEl, "refMode")) calc.refMode = attr(calcPrEl, "refMode");
      if (String(attr(calcPrEl, "calcOnSave")) === "0") calc.calcOnSave = false;
      if (parseOnOff(attr(calcPrEl, "forceFullCalc"))) calc.forceFullCalc = true;
      const cmc = attrNum(calcPrEl, "concurrentManualCount");
      if (cmc !== undefined) calc.concurrentManualCount = cmc;
      if (parseOnOff(attr(calcPrEl, "iterate"))) calc.iterate = true;
      const ic = attrNum(calcPrEl, "iterateCount");
      if (ic !== undefined) calc.iterateCount = ic;
      const id = attrNum(calcPrEl, "iterateDelta");
      if (id !== undefined) calc.iterateDelta = id;
      if (String(attr(calcPrEl, "fullPrecision")) === "0") calc.fullPrecision = false;
      if (parseOnOff(attr(calcPrEl, "calcCompleted"))) calc.calcCompleted = true;
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
        const so = attr(v, "showObjects");
        if (so) view.showObjects = so as CustomWorkbookViewOptions["showObjects"];
        if (String(attr(v, "includeHiddenRowCol")) === "0") view.includeHiddenRowCol = false;
        if (String(attr(v, "includePrintSettings")) === "0") view.includePrintSettings = false;
        if (parseOnOff(attr(v, "personalView"))) view.personalView = true;
        if (parseOnOff(attr(v, "maximized"))) view.maximized = true;
        if (parseOnOff(attr(v, "minimized"))) view.minimized = true;
        if (parseOnOff(attr(v, "autoUpdate"))) view.autoUpdate = true;
        const mi = attrNum(v, "mergeInterval");
        if (mi !== undefined) view.mergeInterval = mi;
        if (parseOnOff(attr(v, "changesSavedWin"))) view.changesSavedWin = true;
        if (parseOnOff(attr(v, "onlySync"))) view.onlySync = true;
        const sc = attr(v, "showComments");
        if (sc) view.showComments = sc as CustomWorkbookViewOptions["showComments"];
        views.push(view);
      }
      if (views.length > 0) result.customViews = views;
    }

    // File sharing
    const fileSharingEl = findChild(el, "fileSharing");
    if (fileSharingEl?.attributes) {
      const fs: FileSharingOptions = {};
      if (parseOnOff(attr(fileSharingEl, "readOnlyRecommended"))) fs.readOnlyRecommended = true;
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
      if (parseOnOff(attr(webPublishingEl, "vml"))) wp.vml = true;
      if (parseOnOff(attr(webPublishingEl, "allowPng"))) wp.allowPng = true;
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
      if (parseOnOff(attr(fileRecoveryEl, "crashSave"))) frp.crashSave = true;
      if (parseOnOff(attr(fileRecoveryEl, "dataExtractLoad"))) frp.dataExtractLoad = true;
      if (parseOnOff(attr(fileRecoveryEl, "repairLoad"))) frp.repairLoad = true;
      result.fileRecoveryPr = frp;
    }

    // Workbook properties
    const wbPrEl = findChild(el, "workbookPr");
    if (wbPrEl?.attributes) {
      const wbPr: WorkbookPropertiesOptions = {};
      if (parseOnOff(attr(wbPrEl, "date1904"))) wbPr.date1904 = true;
      const dtv = attrNum(wbPrEl, "defaultThemeVersion");
      if (dtv !== undefined) wbPr.defaultThemeVersion = dtv;
      if (attr(wbPrEl, "showObjects")) wbPr.showObjects = attr(wbPrEl, "showObjects");
      if (parseOnOff(attr(wbPrEl, "hidePivotFieldList"))) wbPr.hidePivotFieldList = true;
      if (parseOnOff(attr(wbPrEl, "allowRefreshQuery"))) wbPr.allowRefreshQuery = true;
      if (parseOnOff(attr(wbPrEl, "filterPrivacy"))) wbPr.filterPrivacy = true;
      if (parseOnOff(attr(wbPrEl, "backupFile"))) wbPr.backupFile = true;
      if (attr(wbPrEl, "codeName")) wbPr.codeName = attr(wbPrEl, "codeName");
      if (parseOnOff(attr(wbPrEl, "showBorderUnselectedTables")))
        wbPr.showBorderUnselectedTables = true;
      if (parseOnOff(attr(wbPrEl, "promptedSolutions"))) wbPr.promptedSolutions = true;
      if (String(attr(wbPrEl, "showInkAnnotation")) === "0") wbPr.showInkAnnotation = false;
      if (String(attr(wbPrEl, "saveExternalLinkValues")) === "0")
        wbPr.saveExternalLinkValues = false;
      if (attr(wbPrEl, "updateLinks")) wbPr.updateLinks = attr(wbPrEl, "updateLinks");
      if (parseOnOff(attr(wbPrEl, "showPivotChartFilter"))) wbPr.showPivotChartFilter = true;
      if (parseOnOff(attr(wbPrEl, "publishItems"))) wbPr.publishItems = true;
      if (parseOnOff(attr(wbPrEl, "checkCompatibility"))) wbPr.checkCompatibility = true;
      if (String(attr(wbPrEl, "autoCompressPictures")) === "0") wbPr.autoCompressPictures = false;
      if (parseOnOff(attr(wbPrEl, "refreshAllConnections"))) wbPr.refreshAllConnections = true;
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
        if (parseOnOff(attr(wo, "autoRepublish"))) obj.autoRepublish = true;
        if (attr(wo, "title")) obj.title = attr(wo, "title");
        if (attr(wo, "sourceObject")) obj.sourceObject = attr(wo, "sourceObject");
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
        if (parseOnOff(attr(d, "hidden"))) dn.hidden = true;
        if (parseOnOff(attr(d, "function"))) dn.function = true;
        if (parseOnOff(attr(d, "vbProcedure"))) dn.vbProcedure = true;
        if (parseOnOff(attr(d, "xlm"))) dn.xlm = true;
        const functionGroupId = attrNum(d, "functionGroupId");
        if (functionGroupId !== undefined) dn.functionGroupId = functionGroupId;
        if (attr(d, "shortcutKey") !== undefined) dn.shortcutKey = attr(d, "shortcutKey");
        if (parseOnOff(attr(d, "publishToServer"))) dn.publishToServer = true;
        if (parseOnOff(attr(d, "workbookParameter"))) dn.workbookParameter = true;
        names.push(dn);
      }
      if (names.length > 0) result.definedNames = names;
    }

    // Smart tag properties (CT_SmartTagPr)
    const smartTagPrEl = findChild(el, "smartTagPr");
    if (smartTagPrEl?.attributes) {
      const stp: SmartTagPropertiesOptions = {};
      if (parseOnOff(attr(smartTagPrEl, "embed"))) stp.embed = true;
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
