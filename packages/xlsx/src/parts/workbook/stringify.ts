/**
 * Workbook — stringify implementation and exported XML helpers.
 *
 * @module
 */

import { derivePasswordHash } from "@office-open/core";
import { escapeXml } from "@office-open/xml";
import { hashPassword } from "@util/index";

import type { TablePartReference, WorkbookDescriptorOptions } from "./types";

// ── Stringify helpers ──

export function stringifyWorkbook(opts: WorkbookDescriptorOptions): string {
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
  ];

  // File version (CT_Workbook first child). Fresh compiles emit Excel 2007
  // defaults; a round-tripped source carries its own version stamp through.
  const fv = opts.fileVersion;
  if (fv) {
    const fvAttrs: string[] = [];
    if (fv.appName) fvAttrs.push(`appName="${fv.appName}"`);
    if (fv.lastEdited !== undefined) fvAttrs.push(`lastEdited="${fv.lastEdited}"`);
    if (fv.lowestEdited !== undefined) fvAttrs.push(`lowestEdited="${fv.lowestEdited}"`);
    if (fv.rupBuild !== undefined) fvAttrs.push(`rupBuild="${fv.rupBuild}"`);
    parts.push(`<fileVersion ${fvAttrs.join(" ")}/>`);
  } else {
    parts.push('<fileVersion appName="xl" lastEdited="7" lowestEdited="6" rupBuild="29929"/>');
  }

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

  // AbsPath rides in an mc:AlternateContent between workbookPr and bookViews
  if (opts.absPath !== undefined) {
    parts.push(
      '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
        `<mc:Choice Requires="x15"><x15ac:absPath url="${escapeXml(opts.absPath)}" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/></mc:Choice>` +
        "</mc:AlternateContent>",
    );
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
    if (bv.visibility && bv.visibility !== "visible") bvAttrs.push(`visibility="${bv.visibility}"`);
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

  // Defined names (after externalReferences, before calcPr per XSD sequence)
  if (opts.definedNames && opts.definedNames.length > 0) {
    const dnParts: string[] = ["<definedNames>"];
    for (const dn of opts.definedNames) {
      const dnAttrs: string[] = [`name="${escapeXml(dn.name)}"`];
      if (dn.comment !== undefined) dnAttrs.push(`comment="${escapeXml(dn.comment)}"`);
      if (dn.customMenu !== undefined) dnAttrs.push(`customMenu="${escapeXml(dn.customMenu)}"`);
      if (dn.description !== undefined) dnAttrs.push(`description="${escapeXml(dn.description)}"`);
      if (dn.help !== undefined) dnAttrs.push(`help="${escapeXml(dn.help)}"`);
      if (dn.statusBar !== undefined) dnAttrs.push(`statusBar="${escapeXml(dn.statusBar)}"`);
      if (dn.localSheetId !== undefined) dnAttrs.push(`localSheetId="${dn.localSheetId}"`);
      if (dn.hidden) dnAttrs.push('hidden="1"');
      if (dn.function) dnAttrs.push('function="1"');
      if (dn.vbProcedure) dnAttrs.push('vbProcedure="1"');
      if (dn.xlm) dnAttrs.push('xlm="1"');
      if (dn.functionGroupId !== undefined) dnAttrs.push(`functionGroupId="${dn.functionGroupId}"`);
      if (dn.shortcutKey !== undefined) dnAttrs.push(`shortcutKey="${escapeXml(dn.shortcutKey)}"`);
      if (dn.publishToServer) dnAttrs.push('publishToServer="1"');
      if (dn.workbookParameter) dnAttrs.push('workbookParameter="1"');
      dnParts.push(`<definedName ${dnAttrs.join(" ")}>${escapeXml(dn.value)}</definedName>`);
    }
    dnParts.push("</definedNames>");
    parts.push(dnParts.join(""));
  }

  // Calculation properties
  if (opts.calcPr) {
    const cp = opts.calcPr;
    const cpAttrs: string[] = [];
    cpAttrs.push(`calcId="${cp.calcId ?? 191029}"`);
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
    parts.push('<calcPr calcId="191029" fullCalcOnLoad="1"/>');
  }

  // OLE size (after calcPr, before customWorkbookViews per XSD sequence)
  if (opts.oleSize) {
    parts.push(`<oleSize ref="${escapeXml(opts.oleSize)}"/>`);
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
      if (v.showObjects) vAttrs.push(`showObjects="${v.showObjects}"`);
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

  // Smart tag properties (after pivotCaches, before smartTagTypes per XSD sequence)
  if (opts.smartTagPr) {
    const stp = opts.smartTagPr;
    const stpAttrs: string[] = [];
    if (stp.embed) stpAttrs.push('embed="1"');
    if (stp.show && stp.show !== "all") stpAttrs.push(`show="${stp.show}"`);
    if (stpAttrs.length > 0) {
      parts.push(`<smartTagPr ${stpAttrs.join(" ")}/>`);
    }
  }

  // Smart tag types (after smartTagPr, before webPublishing per XSD sequence)
  if (opts.smartTagTypes && opts.smartTagTypes.length > 0) {
    const sttParts: string[] = ["<smartTagTypes>"];
    for (const stt of opts.smartTagTypes) {
      const sttAttrs: string[] = [];
      if (stt.namespaceUri) sttAttrs.push(`namespaceUri="${escapeXml(stt.namespaceUri)}"`);
      if (stt.name) sttAttrs.push(`name="${escapeXml(stt.name)}"`);
      if (stt.url) sttAttrs.push(`url="${escapeXml(stt.url)}"`);
      sttParts.push(`<smartTagType ${sttAttrs.join(" ")}/>`);
    }
    sttParts.push("</smartTagTypes>");
    parts.push(sttParts.join(""));
  }

  // Web publishing (after smartTagTypes, before fileRecoveryPr per XSD sequence)
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

  if (opts.extensions?.length) {
    const exts = opts.extensions
      .map((ext) => {
        const ns = Object.entries(ext.namespaces ?? {})
          .map(([name, value]) => ` ${name}="${escapeXml(value)}"`)
          .join("");
        return `<ext uri="${escapeXml(ext.uri)}"${ns}>${ext.content ?? ""}</ext>`;
      })
      .join("");
    parts.push(`<extLst>${exts}</extLst>`);
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
