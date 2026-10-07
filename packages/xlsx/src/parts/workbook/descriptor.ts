/**
 * Workbook types and descriptor for SpreadsheetML documents.
 *
 * @module
 */

import { parseOnOff, xsdComments } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrNum, findChild, stringifyElement } from "@office-open/xml";

import { stringifyWorkbook } from "./stringify";
import type {
  CalculationPropertiesOptions,
  CustomWorkbookViewOptions,
  DefinedNameOptions,
  FileRecoveryPropertiesOptions,
  FileSharingOptions,
  FileVersionOptions,
  PivotCacheReference,
  RevisionPtrOptions,
  SheetDefinition,
  SmartTagPropertiesOptions,
  SmartTagShow,
  SmartTagTypeOptions,
  WebPublishObjectOptions,
  WebPublishingOptions,
  WorkbookConformance,
  WorkbookDescriptorOptions,
  WorkbookExtensionOptions,
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
    const childNames = (el.elements ?? [])
      .map((child) => child.name?.slice(child.name.indexOf(":") + 1))
      .filter((name): name is string => name !== undefined);
    if (
      childNames.indexOf("bookViews") !== -1 &&
      childNames.indexOf("workbookPr") !== -1 &&
      childNames.indexOf("bookViews") < childNames.indexOf("workbookPr")
    )
      result.legacyChildOrder = true;

    // File version (CT_Workbook first child) — Excel version stamp
    const fileVersionEl = findChild(el, "fileVersion");
    if (fileVersionEl) {
      const fv: FileVersionOptions = {};
      const appName = attr(fileVersionEl, "appName");
      if (appName) fv.appName = appName;
      const codeName = attr(fileVersionEl, "codeName");
      if (codeName) fv.codeName = codeName;
      const lastEdited = attrNum(fileVersionEl, "lastEdited");
      if (lastEdited !== undefined) fv.lastEdited = lastEdited;
      const lowestEdited = attrNum(fileVersionEl, "lowestEdited");
      if (lowestEdited !== undefined) fv.lowestEdited = lowestEdited;
      const rupBuild = attrNum(fileVersionEl, "rupBuild");
      if (rupBuild !== undefined) fv.rupBuild = rupBuild;
      result.fileVersion = fv;
    } else {
      result.fileVersion = false;
    }

    // Sheets
    const sheetsEl = findChild(el, "sheets");
    if (sheetsEl) {
      const sheets: SheetDefinition[] = [];
      for (const s of sheetsEl.elements ?? []) {
        if (s.name !== "sheet") continue;
        const name = attr(s, "name") ?? "";
        const sheetId = attrNum(s, "sheetId");
        const tabId = attrNum(s, "tabId");
        const rId = (s.attributes?.["r:id"] as string | undefined) ?? "";
        const state = attr(s, "state") as SheetDefinition["state"];
        sheets.push({ name, sheetId, tabId, rId, state });
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
    if (protEl) {
      const prot: WorkbookProtectionOptions = {};
      const lockStructure = parseOnOff(attr(protEl, "lockStructure"));
      if (lockStructure !== undefined) prot.lockStructure = lockStructure;
      const lockWindows = parseOnOff(attr(protEl, "lockWindows"));
      if (lockWindows !== undefined) prot.lockWindows = lockWindows;
      const lockRevision = parseOnOff(attr(protEl, "lockRevision"));
      if (lockRevision !== undefined) prot.lockRevision = lockRevision;
      // @workbookPassword/@revisionsPassword (legacy hashes) not read back —
      // the password fields are plaintext authoring input and stringify hashes
      // them, so carrying the hash would double-hash on round-trip. The modern
      // algorithmName/hashValue/saltValue/spinCount quadruplets round-trip as-is.
      if (attr(protEl, "workbookAlgorithmName"))
        prot.workbookAlgorithmName = attr(protEl, "workbookAlgorithmName");
      if (attr(protEl, "workbookHashValue"))
        prot.workbookHashValue = attr(protEl, "workbookHashValue");
      if (attr(protEl, "workbookSaltValue"))
        prot.workbookSaltValue = attr(protEl, "workbookSaltValue");
      if (attr(protEl, "workbookSpinCount"))
        prot.workbookSpinCount = attrNum(protEl, "workbookSpinCount");
      if (attr(protEl, "revisionsAlgorithmName"))
        prot.revisionsAlgorithmName = attr(protEl, "revisionsAlgorithmName");
      if (attr(protEl, "revisionsHashValue"))
        prot.revisionsHashValue = attr(protEl, "revisionsHashValue");
      if (attr(protEl, "revisionsSaltValue"))
        prot.revisionsSaltValue = attr(protEl, "revisionsSaltValue");
      if (attr(protEl, "revisionsSpinCount"))
        prot.revisionsSpinCount = attrNum(protEl, "revisionsSpinCount");
      const wbPasswordHash = attr(protEl, "workbookPassword");
      if (wbPasswordHash !== undefined) prot.workbookPasswordHashRaw = wbPasswordHash;
      const revPasswordHash = attr(protEl, "revisionsPassword");
      if (revPasswordHash !== undefined) prot.revisionsPasswordHashRaw = revPasswordHash;
      if (attr(protEl, "workbookPasswordCharacterSet"))
        prot.workbookPasswordCharacterSet = attr(protEl, "workbookPasswordCharacterSet");
      if (attr(protEl, "revisionsPasswordCharacterSet"))
        prot.revisionsPasswordCharacterSet = attr(protEl, "revisionsPasswordCharacterSet");
      result.protection = prot;
    }

    // Book views
    const bookViewsEl = findChild(el, "bookViews");
    if (bookViewsEl) {
      const bvEl = findChild(bookViewsEl, "workbookView");
      if (bvEl) {
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
        const onOff = (name: string): boolean | undefined => {
          const value = attr(bvEl, name);
          if (value === "0" || value === "false") return false;
          if (value === "1" || value === "true") return true;
          return undefined;
        };
        bv.autoFilterDateGrouping = onOff("autoFilterDateGrouping");
        const fs = attrNum(bvEl, "firstSheet");
        if (fs !== undefined) bv.firstSheet = fs;
        bv.showHorizontalScroll = onOff("showHorizontalScroll");
        bv.showVerticalScroll = onOff("showVerticalScroll");
        bv.showSheetTabs = onOff("showSheetTabs");
        bv.showRowColHeaders = onOff("showRowColHeaders");
        bv.showWhiteSpace = onOff("showWhiteSpace");
        bv.showZeros = onOff("showZeros");
        bv.defaultGridColor = onOff("defaultGridColor");
        bv.windowProtection = onOff("windowProtection");
        bv.rightToLeft = onOff("rightToLeft");
        const minimized = onOff("minimized");
        if (minimized !== undefined) bv.minimized = minimized;
        const tr = attrNum(bvEl, "tabRatio");
        if (tr !== undefined) bv.tabRatio = tr;
        const vis = attr(bvEl, "visibility");
        if (vis !== undefined) bv.visibility = vis as WorkbookViewOptions["visibility"];
        const uid = attr(bvEl, "xr2:uid");
        if (uid !== undefined) bv.uid = uid;
        result.bookView = bv;
      }
    } else {
      result.bookView = false;
    }

    // Calc properties
    const calcPrEl = findChild(el, "calcPr");
    if (calcPrEl) {
      const calc: CalculationPropertiesOptions = {};
      const calcId = attrNum(calcPrEl, "calcId");
      if (calcId !== undefined) calc.calcId = calcId;
      if (attr(calcPrEl, "calcMode"))
        calc.calcMode = attr(calcPrEl, "calcMode") as CalculationPropertiesOptions["calcMode"];
      const fullCalcOnLoad = parseOnOff(attr(calcPrEl, "fullCalcOnLoad"));
      if (fullCalcOnLoad !== undefined) calc.fullCalcOnLoad = fullCalcOnLoad;
      const concurrentCalc = parseOnOff(attr(calcPrEl, "concurrentCalc"));
      if (concurrentCalc !== undefined) calc.concurrentCalc = concurrentCalc;
      if (attr(calcPrEl, "refMode"))
        calc.refMode = attr(calcPrEl, "refMode") as CalculationPropertiesOptions["refMode"];
      const calcOnSave = parseOnOff(attr(calcPrEl, "calcOnSave"));
      if (calcOnSave !== undefined) calc.calcOnSave = calcOnSave;
      const forceFullCalc = parseOnOff(attr(calcPrEl, "forceFullCalc"));
      if (forceFullCalc !== undefined) calc.forceFullCalc = forceFullCalc;
      const cmc = attrNum(calcPrEl, "concurrentManualCount");
      if (cmc !== undefined) calc.concurrentManualCount = cmc;
      const iterate = parseOnOff(attr(calcPrEl, "iterate"));
      if (iterate !== undefined) calc.iterate = iterate;
      const ic = attrNum(calcPrEl, "iterateCount");
      if (ic !== undefined) calc.iterateCount = ic;
      const id = attrNum(calcPrEl, "iterateDelta");
      if (id !== undefined) calc.iterateDelta = id;
      const fullPrecision = parseOnOff(attr(calcPrEl, "fullPrecision"));
      if (fullPrecision !== undefined) calc.fullPrecision = fullPrecision;
      if (attr(calcPrEl, "calcCompleted") !== undefined)
        calc.calcCompleted = parseOnOff(attr(calcPrEl, "calcCompleted")) ?? false;
      result.calculation = calc;
    }

    // OLE size (after calcPr, before customWorkbookViews per XSD sequence)
    const oleSizeEl = findChild(el, "oleSize");
    if (oleSizeEl) {
      const ref = attr(oleSizeEl, "ref");
      if (ref) result.oleSize = ref;
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
        if (sc)
          view.showComments = xsdComments.from(sc) as CustomWorkbookViewOptions["showComments"];
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
    if (webPublishingEl) {
      const wp: WebPublishingOptions = {};
      if (String(attr(webPublishingEl, "css")) === "0") wp.css = false;
      if (String(attr(webPublishingEl, "thicket")) === "0") wp.thicket = false;
      if (String(attr(webPublishingEl, "longFileNames")) === "0") wp.longFileNames = false;
      if (parseOnOff(attr(webPublishingEl, "vml"))) wp.vml = true;
      if (parseOnOff(attr(webPublishingEl, "allowPng"))) wp.allowPng = true;
      if (attr(webPublishingEl, "targetScreenSize"))
        wp.targetScreenSize = attr(
          webPublishingEl,
          "targetScreenSize",
        ) as WebPublishingOptions["targetScreenSize"];
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
      if (attr(fileRecoveryEl, "autoRecover") !== undefined)
        frp.autoRecover = parseOnOff(attr(fileRecoveryEl, "autoRecover")) ?? false;
      if (parseOnOff(attr(fileRecoveryEl, "crashSave"))) frp.crashSave = true;
      if (parseOnOff(attr(fileRecoveryEl, "dataExtractLoad"))) frp.dataExtractLoad = true;
      if (parseOnOff(attr(fileRecoveryEl, "repairLoad"))) frp.repairLoad = true;
      result.fileRecovery = frp;
    }

    // Workbook properties
    const wbPrEl = findChild(el, "workbookPr");
    if (wbPrEl) {
      const wbPr: WorkbookPropertiesOptions = {};
      if (attr(wbPrEl, "date1904") !== undefined)
        wbPr.date1904 = parseOnOff(attr(wbPrEl, "date1904"));
      if (attr(wbPrEl, "dateCompatibility") !== undefined)
        wbPr.dateCompatibility = parseOnOff(attr(wbPrEl, "dateCompatibility"));
      const dtv = attrNum(wbPrEl, "defaultThemeVersion");
      if (dtv !== undefined) wbPr.defaultThemeVersion = dtv;
      if (attr(wbPrEl, "showObjects"))
        wbPr.showObjects = attr(wbPrEl, "showObjects") as WorkbookPropertiesOptions["showObjects"];
      const hidePivotFieldList = parseOnOff(attr(wbPrEl, "hidePivotFieldList"));
      if (hidePivotFieldList !== undefined) wbPr.hidePivotFieldList = hidePivotFieldList;
      const allowRefreshQuery = parseOnOff(attr(wbPrEl, "allowRefreshQuery"));
      if (allowRefreshQuery !== undefined) wbPr.allowRefreshQuery = allowRefreshQuery;
      const filterPrivacy = parseOnOff(attr(wbPrEl, "filterPrivacy"));
      if (filterPrivacy !== undefined) wbPr.filterPrivacy = filterPrivacy;
      if (attr(wbPrEl, "backupFile") !== undefined)
        wbPr.backupFile = parseOnOff(attr(wbPrEl, "backupFile"));
      if (attr(wbPrEl, "codeName")) wbPr.codeName = attr(wbPrEl, "codeName");
      const showBorderUnselectedTables = parseOnOff(attr(wbPrEl, "showBorderUnselectedTables"));
      if (showBorderUnselectedTables !== undefined)
        wbPr.showBorderUnselectedTables = showBorderUnselectedTables;
      const promptedSolutions = parseOnOff(attr(wbPrEl, "promptedSolutions"));
      if (promptedSolutions !== undefined) wbPr.promptedSolutions = promptedSolutions;
      const showInkAnnotation = parseOnOff(attr(wbPrEl, "showInkAnnotation"));
      if (showInkAnnotation !== undefined) wbPr.showInkAnnotation = showInkAnnotation;
      const saveExternalLinkValues = parseOnOff(attr(wbPrEl, "saveExternalLinkValues"));
      if (saveExternalLinkValues !== undefined)
        wbPr.saveExternalLinkValues = saveExternalLinkValues;
      if (attr(wbPrEl, "updateLinks"))
        wbPr.updateLinks = attr(wbPrEl, "updateLinks") as WorkbookPropertiesOptions["updateLinks"];
      const showPivotChartFilter = parseOnOff(attr(wbPrEl, "showPivotChartFilter"));
      if (showPivotChartFilter !== undefined) wbPr.showPivotChartFilter = showPivotChartFilter;
      const publishItems = parseOnOff(attr(wbPrEl, "publishItems"));
      if (publishItems !== undefined) wbPr.publishItems = publishItems;
      const checkCompatibility = parseOnOff(attr(wbPrEl, "checkCompatibility"));
      if (checkCompatibility !== undefined) wbPr.checkCompatibility = checkCompatibility;
      const autoCompressPictures = parseOnOff(attr(wbPrEl, "autoCompressPictures"));
      if (autoCompressPictures !== undefined) wbPr.autoCompressPictures = autoCompressPictures;
      const refreshAllConnections = parseOnOff(attr(wbPrEl, "refreshAllConnections"));
      if (refreshAllConnections !== undefined) wbPr.refreshAllConnections = refreshAllConnections;
      result.properties = wbPr;
    }

    // AbsPath rides in an mc:AlternateContent between workbookPr and bookViews
    const acEl = findChild(el, "mc:AlternateContent");
    if (acEl) {
      const choice = findChild(acEl, "mc:Choice");
      // Excel writes both prefixes in the wild: x15ac (2010/11/ac, the
      // common form our stringify re-emits) and x15 (2010/11/main, root-bound).
      const legacyAbsPathEl = choice ? findChild(choice, "x15ac:absPath") : undefined;
      const absPathEl = legacyAbsPathEl ?? (choice ? findChild(choice, "x15:absPath") : undefined);
      const url = absPathEl ? attr(absPathEl, "url") : undefined;
      if (url !== undefined) result.absPath = url;
      if (legacyAbsPathEl) result.absPathLegacyPrefix = true;
    }

    // Coauthoring revision state
    const rpEl = findChild(el, "xr:revisionPtr");
    if (rpEl) {
      const rp: RevisionPtrOptions = {};
      const revId = attrNum(rpEl, "revIDLastSave");
      if (revId !== undefined) rp.revisionIdLastSave = revId;
      if (attr(rpEl, "documentId")) rp.documentId = attr(rpEl, "documentId");
      const cvLast = attrNum(rpEl, "xr6:coauthVersionLast");
      if (cvLast !== undefined) rp.coauthVersionLast = cvLast;
      const cvMax = attrNum(rpEl, "xr6:coauthVersionMax");
      if (cvMax !== undefined) rp.coauthVersionMax = cvMax;
      if (attr(rpEl, "xr10:uidLastSave")) rp.uidLastSave = attr(rpEl, "xr10:uidLastSave");
      result.revisionPtr = rp;
    }

    // Trailing extension list (workbook > extLst > ext) — raw round-trip
    const extLstEl = findChild(el, "extLst");
    if (extLstEl) {
      const extensions: WorkbookExtensionOptions[] = [];
      for (const ext of extLstEl.elements ?? []) {
        if (ext.name !== "ext") continue;
        const uri = attr(ext, "uri");
        if (!uri) continue;
        const namespaces: Record<string, string> = {};
        for (const [name, value] of Object.entries(ext.attributes ?? {})) {
          if (name.startsWith("xmlns:") && typeof value === "string") namespaces[name] = value;
        }
        const content = (ext.elements ?? []).map((e) => stringifyElement(e)).join("");
        extensions.push({
          uri,
          ...(Object.keys(namespaces).length > 0 ? { namespaces } : {}),
          ...(content ? { content } : {}),
        });
      }
      if (extensions.length > 0) result.extensions = extensions;
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
      result.functionGroups = names;
    }

    // Web publish objects
    const wpoEl = findChild(el, "webPublishObjects");
    if (wpoEl) {
      const objs: WebPublishObjectOptions[] = [];
      for (const wo of wpoEl.elements ?? []) {
        if (wo.name !== "webPublishObject") continue;
        const id = attrNum(wo, "id");
        const divId = attr(wo, "divId");
        const destinationFile = attr(wo, "destinationFile");
        if (id === undefined || divId === undefined || destinationFile === undefined) continue;
        const obj: WebPublishObjectOptions = { id, divId, destinationFile };
        if (attr(wo, "sourceObject")) obj.sourceObject = attr(wo, "sourceObject");
        if (attr(wo, "title")) obj.title = attr(wo, "title");
        if (parseOnOff(attr(wo, "autoRepublish"))) obj.autoRepublish = true;
        objs.push(obj);
      }
      if (objs.length > 0) result.webPublishObjects = objs;
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
      result.definedNames = names;
    }

    // Smart tag properties (CT_SmartTagPr)
    const smartTagPrEl = findChild(el, "smartTagPr");
    if (smartTagPrEl?.attributes) {
      const stp: SmartTagPropertiesOptions = {};
      if (parseOnOff(attr(smartTagPrEl, "embed"))) stp.embed = true;
      const show = attr(smartTagPrEl, "show");
      if (show && show !== "all") stp.show = show as SmartTagShow;
      result.smartTag = stp;
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
