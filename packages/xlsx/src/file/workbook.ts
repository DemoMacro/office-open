/**
 * Workbook types for SpreadsheetML documents.
 *
 * @module
 */

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
