/**
 * Workbook component — generates xl/workbook.xml.
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { derivePasswordHash } from "@office-open/core";
import { escapeXml } from "@office-open/xml";

export interface SheetDefinition {
  readonly name: string;
  readonly sheetId: number;
  readonly rId: string;
  readonly state?: "visible" | "hidden" | "veryHidden";
}

export interface PivotCacheReference {
  readonly cacheId: number;
  readonly rId: string;
}

export interface TablePartReference {
  readonly rId: string;
}

/** Custom workbook view for storing display preferences. */
export interface CustomWorkbookViewOptions {
  /** View name */
  readonly name: string;
  /** GUID (e.g. "{00000000-0000-0000-0000-000000000000}") */
  readonly guid: string;
  /** Window width in twips */
  readonly windowWidth: number;
  /** Window height in twips */
  readonly windowHeight: number;
  /** Active sheet ID (1-based sheetId) */
  readonly activeSheetId: number;
  /** X position of the window */
  readonly xWindow?: number;
  /** Y position of the window */
  readonly yWindow?: number;
  /** Show formula bar (default true) */
  readonly showFormulaBar?: boolean;
  /** Show status bar (default true) */
  readonly showStatusbar?: boolean;
  /** Show horizontal scroll (default true) */
  readonly showHorizontalScroll?: boolean;
  /** Show vertical scroll (default true) */
  readonly showVerticalScroll?: boolean;
  /** Show sheet tabs (default true) */
  readonly showSheetTabs?: boolean;
  /** Tab ratio (default 600) */
  readonly tabRatio?: number;
  /** Include hidden rows/columns (default true) */
  readonly includeHiddenRowCol?: boolean;
  /** Include print settings (default true) */
  readonly includePrintSettings?: boolean;
  /** Personal view (default false) */
  readonly personalView?: boolean;
  /** Maximized (default false) */
  readonly maximized?: boolean;
  /** Minimized (default false) */
  readonly minimized?: boolean;
  /** Auto update (CT_CustomWorkbookView @autoUpdate) */
  readonly autoUpdate?: boolean;
  /** Merge interval (CT_CustomWorkbookView @mergeInterval) */
  readonly mergeInterval?: number;
  /** Changes saved in window (CT_CustomWorkbookView @changesSavedWin) */
  readonly changesSavedWin?: boolean;
  /** Only sync (CT_CustomWorkbookView @onlySync) */
  readonly onlySync?: boolean;
  /** Show comments (CT_CustomWorkbookView @showComments) */
  readonly showComments?: string;
}

export interface WorkbookProtectionOptions {
  /** Lock workbook structure (add/delete/rename/move sheets) */
  readonly lockStructure?: boolean;
  /** Lock workbook windows */
  readonly lockWindows?: boolean;
  /** Lock revisions */
  readonly lockRevision?: boolean;
  /** Plain-text password — legacy Excel hash computed automatically */
  readonly workbookPassword?: string;
  /** Modern encryption: algorithm name (e.g. "SHA-512") */
  readonly workbookAlgorithmName?: string;
  /** Modern encryption: base64-encoded hash value */
  readonly workbookHashValue?: string;
  /** Modern encryption: base64-encoded salt value */
  readonly workbookSaltValue?: string;
  /** Modern encryption: spin count */
  readonly workbookSpinCount?: number;
  /** Revisions password (legacy) */
  readonly revisionsPassword?: string;
  /** Revisions modern encryption: algorithm name */
  readonly revisionsAlgorithmName?: string;
  /** Revisions modern encryption: base64-encoded hash value */
  readonly revisionsHashValue?: string;
  /** Revisions modern encryption: base64-encoded salt value */
  readonly revisionsSaltValue?: string;
  /** Revisions modern encryption: spin count */
  readonly revisionsSpinCount?: number;
}

/** File recovery properties (CT_FileRecoveryPr) */
export interface FileRecoveryPrOptions {
  /** Enable auto-recover (default true) */
  readonly autoRecover?: boolean;
  /** Crash save (default false) */
  readonly crashSave?: boolean;
  /** Data extract load (default false) */
  readonly dataExtractLoad?: boolean;
  /** Repair load (default false) */
  readonly repairLoad?: boolean;
}

/** Web publishing properties (CT_WebPublishing) */
export interface WebPublishingOptions {
  /** Use CSS (default true) */
  readonly css?: boolean;
  /** Use thicket format (default true) */
  readonly thicket?: boolean;
  /** Use long file names (default true) */
  readonly longFileNames?: boolean;
  /** Use VML (default false) */
  readonly vml?: boolean;
  /** Allow PNG (default false) */
  readonly allowPng?: boolean;
  /** Target screen size (default "800x600") */
  readonly targetScreenSize?: string;
  /** DPI (default 96) */
  readonly dpi?: number;
  /** Code page */
  readonly codePage?: number;
  /** Character set */
  readonly characterSet?: string;
}

/** File sharing properties (CT_FileSharing) */
export interface FileSharingOptions {
  /** Recommend read-only mode (default false) */
  readonly readOnlyRecommended?: boolean;
  /** User name who has the file locked */
  readonly userName?: string;
  /** Legacy reservation password (hex) */
  readonly reservationPassword?: string;
  /** Modern encryption: algorithm name */
  readonly algorithmName?: string;
  /** Modern encryption: base64 hash value */
  readonly hashValue?: string;
  /** Modern encryption: base64 salt value */
  readonly saltValue?: string;
  /** Modern encryption: spin count */
  readonly spinCount?: number;
}

/** Workbook properties (CT_WorkbookPr) */
export interface WorkbookPrOptions {
  /** Use 1904 date system (default false) */
  readonly date1904?: boolean;
  /** Default theme version */
  readonly defaultThemeVersion?: number;
  /** Show objects: "all" | "placeholders" | "none" */
  readonly showObjects?: string;
  /** Hide pivot field list (default false) */
  readonly hidePivotFieldList?: boolean;
  /** Allow refresh queries (default false) */
  readonly allowRefreshQuery?: boolean;
  /** Filter privacy (default false) */
  readonly filterPrivacy?: boolean;
  /** Backup file (default false) */
  readonly backupFile?: boolean;
  /** Code name */
  readonly codeName?: string;
  /** Show border unselected tables (CT_WorkbookPr @showBorderUnselectedTables) */
  readonly showBorderUnselectedTables?: boolean;
  /** Prompted solutions (CT_WorkbookPr @promptedSolutions) */
  readonly promptedSolutions?: boolean;
  /** Show ink annotation (CT_WorkbookPr @showInkAnnotation) */
  readonly showInkAnnotation?: boolean;
  /** Save external link values (CT_WorkbookPr @saveExternalLinkValues) */
  readonly saveExternalLinkValues?: boolean;
  /** Update links mode (CT_WorkbookPr @updateLinks) */
  readonly updateLinks?: string;
  /** Show pivot chart filter (CT_WorkbookPr @showPivotChartFilter) */
  readonly showPivotChartFilter?: boolean;
  /** Publish items (CT_WorkbookPr @publishItems) */
  readonly publishItems?: boolean;
  /** Check compatibility (CT_WorkbookPr @checkCompatibility) */
  readonly checkCompatibility?: boolean;
  /** Auto compress pictures (CT_WorkbookPr @autoCompressPictures) */
  readonly autoCompressPictures?: boolean;
  /** Refresh all connections (CT_WorkbookPr @refreshAllConnections) */
  readonly refreshAllConnections?: boolean;
}

/** Volatile type entry (CT_VolType) */
export interface VolTypeOptions {
  /** Type of volatile dependency (default: "realTimeData") */
  readonly type?: "realTimeData" | "olapFunctions";
  /** Main volatile dependencies (CT_VolMain, required) */
  readonly mains?: readonly VolMainOptions[];
}

/** Main volatile dependency (CT_VolMain) */
export interface VolMainOptions {
  /** First reference (required) */
  readonly first: string;
  /** Volatile topics (CT_VolTopic) */
  readonly topics?: readonly VolTopicOptions[];
}

/** Volatile topic (CT_VolTopic) */
export interface VolTopicOptions {
  /** Topic value (required) */
  readonly value: string;
  /** Value type (default: "n") */
  readonly valueType?: string;
  /** String topics (stp elements) */
  readonly stringTopics?: readonly string[];
  /** Topic references (CT_VolTopicRef) */
  readonly refs?: readonly VolTopicRefOptions[];
}

/** Volatile topic reference (CT_VolTopicRef) */
export interface VolTopicRefOptions {
  /** Cell reference (required) */
  readonly reference: string;
  /** Sheet index (required) */
  readonly sheetIndex: number;
}

/** Web publish object (CT_WebPublishObject) */
export interface WebPublishObjectOptions {
  /** Relationship ID to the published item */
  readonly rId: string;
  /** Destination file name */
  readonly destinationFile?: string;
  /** Auto republish (default: false) */
  readonly autoRepublish?: boolean;
  /** Title of the published item */
  readonly title?: string;
  /** Source object reference */
  readonly sourceObject?: string;
  /** App name (default: "Excel") */
  readonly appName?: string;
}

/** Calculation properties (CT_CalcPr) */
export interface CalcPrOptions {
  /** Calculation mode: "manual" | "auto" | "autoNoTable" */
  readonly calcMode?: string;
  /** Calc ID (default 162913) */
  readonly calcId?: number;
  /** Full calc on load (default false) */
  readonly fullCalcOnLoad?: boolean;
  /** Calc on save (default true) */
  readonly calcOnSave?: boolean;
  /** Force full calc */
  readonly forceFullCalc?: boolean;
  /** Concurrent calc (default true) */
  readonly concurrentCalc?: boolean;
  /** Concurrent manual count */
  readonly concurrentManualCount?: number;
  /** Iterate (default false) */
  readonly iterate?: boolean;
  /** Iterate count (default 100) */
  readonly iterateCount?: number;
  /** Iterate delta (default 0.001) */
  readonly iterateDelta?: number;
  /** Reference mode: "A1" | "R1C1" */
  readonly refMode?: string;
  /** Full precision (default true) */
  readonly fullPrecision?: boolean;
  /** Calc completed (CT_CalcPr @calcCompleted) */
  readonly calcCompleted?: boolean;
}

/** Workbook view options (CT_BookView) */
export interface WorkbookViewOptions {
  /** Active tab index (0-based) */
  readonly activeTab?: number;
  /** Auto filter date grouping (default true) */
  readonly autoFilterDateGrouping?: boolean;
  /** First sheet tab */
  readonly firstSheet?: number;
  /** Show horizontal scroll (default true) */
  readonly showHorizontalScroll?: boolean;
  /** Show sheet tabs (default true) */
  readonly showSheetTabs?: boolean;
  /** Show vertical scroll (default true) */
  readonly showVerticalScroll?: boolean;
  /** Tab ratio (default 600) */
  readonly tabRatio?: number;
  /** Window width in twips */
  readonly windowWidth?: number;
  /** Window height in twips */
  readonly windowHeight?: number;
  /** X position of the window */
  readonly xWindow?: number;
  /** Y position of the window */
  readonly yWindow?: number;
}

export class WorkbookXml extends BaseXmlComponent {
  private readonly sheets: readonly SheetDefinition[];
  private readonly pivotCaches: readonly PivotCacheReference[];
  private readonly protection?: WorkbookProtectionOptions;
  private readonly customViews?: readonly CustomWorkbookViewOptions[];
  private readonly fileRecoveryPr?: FileRecoveryPrOptions;
  private readonly functionGroupNames: readonly string[];
  private readonly webPublishing?: WebPublishingOptions;
  private readonly fileSharing?: FileSharingOptions;
  private readonly workbookPr?: WorkbookPrOptions;
  private readonly calcPr?: CalcPrOptions;
  private readonly bookView?: WorkbookViewOptions;
  private readonly volTypes?: readonly VolTypeOptions[];
  private readonly webPublishObjects?: readonly WebPublishObjectOptions[];

  public constructor(
    sheets: readonly SheetDefinition[],
    pivotCaches?: readonly PivotCacheReference[],
    protection?: WorkbookProtectionOptions,
    customViews?: readonly CustomWorkbookViewOptions[],
    fileRecoveryPr?: FileRecoveryPrOptions,
    functionGroups?: readonly string[],
    webPublishing?: WebPublishingOptions,
    fileSharing?: FileSharingOptions,
    workbookPr?: WorkbookPrOptions,
    calcPr?: CalcPrOptions,
    bookView?: WorkbookViewOptions,
    volTypes?: readonly VolTypeOptions[],
    webPublishObjects?: readonly WebPublishObjectOptions[],
  ) {
    super("workbook");
    this.sheets = sheets;
    this.pivotCaches = pivotCaches ?? [];
    this.protection = protection;
    this.customViews = customViews;
    this.fileRecoveryPr = fileRecoveryPr;
    this.functionGroupNames = functionGroups ?? [];
    this.webPublishing = webPublishing;
    this.fileSharing = fileSharing;
    this.workbookPr = workbookPr;
    this.calcPr = calcPr;
    this.bookView = bookView;
    this.volTypes = volTypes;
    this.webPublishObjects = webPublishObjects;
  }

  public override toXml(_context: Context): string {
    const parts: string[] = [
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
        ' mc:Ignorable="x15 xr xr6 xr10 xr2"' +
        ' xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"' +
        ' xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"' +
        ' xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6"' +
        ' xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10"' +
        ' xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2">',
      '<fileVersion appName="xl" lastEdited="7" lowestEdited="6" rupBuild="29929"/>',
    ];

    // File sharing (after fileVersion, before workbookPr per XSD sequence)
    if (this.fileSharing) {
      const fileSharing = this.fileSharing;
      const fsAttrs: string[] = [];
      if (fileSharing.readOnlyRecommended) fsAttrs.push('readOnlyRecommended="1"');
      if (fileSharing.userName) fsAttrs.push(`userName="${escapeXml(fileSharing.userName)}"`);
      if (fileSharing.reservationPassword) {
        fsAttrs.push(`reservationPassword="${escapeXml(fileSharing.reservationPassword)}"`);
        // Auto-derive modern hash when reservationPassword provided without explicit hashValue
        if (fileSharing.hashValue === undefined) {
          const derived = derivePasswordHash(fileSharing.reservationPassword);
          fsAttrs.push(`algorithmName="${escapeXml(derived.algorithmName)}"`);
          fsAttrs.push(`hashValue="${escapeXml(derived.hashValue)}"`);
          fsAttrs.push(`saltValue="${escapeXml(derived.saltValue)}"`);
          fsAttrs.push(`spinCount="${derived.spinCount}"`);
        }
      }
      if (fileSharing.algorithmName)
        fsAttrs.push(`algorithmName="${escapeXml(fileSharing.algorithmName)}"`);
      if (fileSharing.hashValue) fsAttrs.push(`hashValue="${escapeXml(fileSharing.hashValue)}"`);
      if (fileSharing.saltValue) fsAttrs.push(`saltValue="${escapeXml(fileSharing.saltValue)}"`);
      if (fileSharing.spinCount !== undefined) fsAttrs.push(`spinCount="${fileSharing.spinCount}"`);
      if (fsAttrs.length > 0) {
        parts.push(`<fileSharing ${fsAttrs.join(" ")}/>`);
      }
    }

    // Workbook properties
    if (this.workbookPr) {
      const wbPr = this.workbookPr;
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
    if (this.protection) {
      const prot = this.protection;
      const protAttrs: string[] = [];
      if (prot.lockStructure) protAttrs.push('lockStructure="1"');
      if (prot.lockWindows) protAttrs.push('lockWindows="1"');
      if (prot.lockRevision) protAttrs.push('lockRevision="1"');
      if (prot.workbookPassword) {
        protAttrs.push(`workbookPassword="${this.hashPassword(prot.workbookPassword)}"`);
        // Auto-derive modern hash when workbookPassword provided without explicit workbookHashValue
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
        protAttrs.push(`revisionsPassword="${this.hashPassword(prot.revisionsPassword)}"`);
        // Auto-derive modern hash when revisionsPassword provided without explicit revisionsHashValue
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
      if (protAttrs.length > 0) {
        parts.push(`<workbookProtection ${protAttrs.join(" ")}/>`);
      }
    }

    // Book views
    if (this.bookView) {
      const bv = this.bookView;
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
    for (const s of this.sheets) {
      const stateAttr = s.state && s.state !== "visible" ? ` state="${s.state}"` : "";
      parts.push(
        `<sheet name="${escapeXml(s.name)}" sheetId="${s.sheetId}" r:id="${s.rId}"${stateAttr}/>`,
      );
    }
    parts.push("</sheets>");

    // Function groups (after sheets, before externalReferences per XSD)
    if (this.functionGroupNames.length > 0) {
      const functionGroupParts: string[] = [`<functionGroups builtInGroupCount="16">`];
      for (const name of this.functionGroupNames) {
        functionGroupParts.push(`<functionGroup name="${escapeXml(name)}"/>`);
      }
      functionGroupParts.push("</functionGroups>");
      parts.push(functionGroupParts.join(""));
    }

    // externalReferences placeholder — compiler injects the XML here if needed
    parts.push("<!--EXTERNAL_REFS-->");

    // Calculation properties
    if (this.calcPr) {
      const cp = this.calcPr;
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
    if (this.customViews && this.customViews.length > 0) {
      parts.push("<customWorkbookViews>");
      for (const v of this.customViews) {
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

    if (this.pivotCaches.length > 0) {
      parts.push("<pivotCaches>");
      for (const pc of this.pivotCaches) {
        parts.push(`<pivotCache cacheId="${pc.cacheId}" r:id="${pc.rId}"/>`);
      }
      parts.push("</pivotCaches>");
    }

    // Web publishing (after pivotCaches, before fileRecoveryPr per XSD sequence)
    if (this.webPublishing) {
      const webPublishing = this.webPublishing;
      const wpAttrs: string[] = [];
      if (webPublishing.css === false) wpAttrs.push('css="0"');
      if (webPublishing.thicket === false) wpAttrs.push('thicket="0"');
      if (webPublishing.longFileNames === false) wpAttrs.push('longFileNames="0"');
      if (webPublishing.vml) wpAttrs.push('vml="1"');
      if (webPublishing.allowPng) wpAttrs.push('allowPng="1"');
      if (webPublishing.targetScreenSize && webPublishing.targetScreenSize !== "800x600")
        wpAttrs.push(`targetScreenSize="${webPublishing.targetScreenSize}"`);
      if (webPublishing.dpi !== undefined && webPublishing.dpi !== 96)
        wpAttrs.push(`dpi="${webPublishing.dpi}"`);
      if (webPublishing.codePage !== undefined)
        wpAttrs.push(`codePage="${webPublishing.codePage}"`);
      if (webPublishing.characterSet)
        wpAttrs.push(`characterSet="${escapeXml(webPublishing.characterSet)}"`);
      parts.push(`<webPublishing ${wpAttrs.join(" ")}/>`);
    }

    // File recovery properties (after webPublishing per XSD sequence)
    if (this.fileRecoveryPr) {
      const fileRecovery = this.fileRecoveryPr;
      const frpAttrs: string[] = [];
      if (fileRecovery.autoRecover === false) frpAttrs.push('autoRecover="0"');
      if (fileRecovery.crashSave) frpAttrs.push('crashSave="1"');
      if (fileRecovery.dataExtractLoad) frpAttrs.push('dataExtractLoad="1"');
      if (fileRecovery.repairLoad) frpAttrs.push('repairLoad="1"');
      if (frpAttrs.length > 0) {
        parts.push(`<fileRecoveryPr ${frpAttrs.join(" ")}/>`);
      }
    }

    // Web publish objects (after fileRecoveryPr per XSD sequence)
    if (this.webPublishObjects && this.webPublishObjects.length > 0) {
      const wpoParts: string[] = [`<webPublishObjects count="${this.webPublishObjects.length}">`];
      for (const wpo of this.webPublishObjects) {
        const wpoAttrs: string[] = [`r:id="${escapeXml(wpo.rId)}"`];
        if (wpo.destinationFile)
          wpoAttrs.push(`destinationFile="${escapeXml(wpo.destinationFile)}"`);
        if (wpo.autoRepublish) wpoAttrs.push('autoRepublish="1"');
        if (wpo.title) wpoAttrs.push(`title="${escapeXml(wpo.title)}"`);
        if (wpo.sourceObject) wpoAttrs.push(`sourceObject="${escapeXml(wpo.sourceObject)}"`);
        wpoParts.push(`<webPublishObject ${wpoAttrs.join(" ")}/>`);
      }
      wpoParts.push("</webPublishObjects>");
      parts.push(wpoParts.join(""));
    }

    // Volatile dependencies (volTypes)
    if (this.volTypes && this.volTypes.length > 0) {
      const vtParts: string[] = [`<volTypes count="${this.volTypes.length}">`];
      for (const vt of this.volTypes) {
        const vtType = vt.type ?? "realTimeData";
        const mains = vt.mains ?? [];
        if (mains.length > 0) {
          const mainParts: string[] = [];
          for (const m of mains) {
            const tpParts: string[] = [];
            for (const topic of m.topics ?? []) {
              let tpInner = `<v>${escapeXml(topic.value)}</v>`;
              for (const stp of topic.stringTopics ?? []) {
                tpInner += `<stp>${escapeXml(stp)}</stp>`;
              }
              for (const tr of topic.refs ?? []) {
                tpInner += `<tr r="${escapeXml(tr.reference)}" s="${tr.sheetIndex}"/>`;
              }
              const tpAttr =
                topic.valueType && topic.valueType !== "n"
                  ? ` t="${escapeXml(topic.valueType)}"`
                  : "";
              tpParts.push(`<tp${tpAttr}>${tpInner}</tp>`);
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

  /**
   * Generate tableParts XML fragment for embedding in a worksheet.
   * This is called by the compiler to insert table references into the worksheet XML.
   */
  public static buildTablePartsXml(tableParts: readonly TablePartReference[]): string {
    if (tableParts.length === 0) return "";
    const parts: string[] = [`<tableParts count="${tableParts.length}">`];
    for (const tp of tableParts) {
      parts.push(`<tablePart r:id="${tp.rId}"/>`);
    }
    parts.push("</tableParts>");
    return parts.join("");
  }

  /**
   * Generate externalReferences XML fragment for embedding in the workbook.
   * This is called by the compiler to insert external reference entries.
   */
  public static buildExternalReferencesXml(refs: readonly { rId: string }[]): string {
    if (refs.length === 0) return "";
    const parts: string[] = ["<externalReferences>"];
    for (const ref of refs) {
      parts.push(`<externalReference r:id="${ref.rId}"/>`);
    }
    parts.push("</externalReferences>");
    return parts.join("");
  }

  /** Legacy Excel password hash (XOR-based) */
  private hashPassword(password: string): string {
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
}
