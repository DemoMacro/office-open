/**
 * Workbook — option types for SpreadsheetML documents.
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
  /** Object display mode (CT_CustomWorkbookView `@showObjects`, ST_Objects) */
  showObjects?: "all" | "placeholders" | "none";
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
  /** Auto update (CT_CustomWorkbookView `@autoUpdate`) */
  autoUpdate?: boolean;
  /** Merge interval (CT_CustomWorkbookView `@mergeInterval`) */
  mergeInterval?: number;
  /** Changes saved in window (CT_CustomWorkbookView `@changesSavedWin`) */
  changesSavedWin?: boolean;
  /** Only sync (CT_CustomWorkbookView `@onlySync`) */
  onlySync?: boolean;
  /** Show comments (CT_CustomWorkbookView `@showComments`, ST_Comments) */
  showComments?: "none" | "indicator" | "comment";
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
  /** Workbook password character set (CT_WorkbookProtection `@workbookPasswordCharacterSet`) */
  workbookPasswordCharacterSet?: string;
  /** Revisions password character set (CT_WorkbookProtection `@revisionsPasswordCharacterSet`) */
  revisionsPasswordCharacterSet?: string;
}

/** Workbook conformance level (CT_Workbook `@conformance`) */
export type WorkbookConformance = "strict" | "transitional";

/** File recovery properties (CT_FileRecoveryPr) */
export interface FileRecoveryPropertiesOptions {
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
export interface WorkbookPropertiesOptions {
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
  /** Show border unselected tables (CT_WorkbookPr `@showBorderUnselectedTables`) */
  showBorderUnselectedTables?: boolean;
  /** Prompted solutions (CT_WorkbookPr `@promptedSolutions`) */
  promptedSolutions?: boolean;
  /** Show ink annotation (CT_WorkbookPr `@showInkAnnotation`) */
  showInkAnnotation?: boolean;
  /** Save external link values (CT_WorkbookPr `@saveExternalLinkValues`) */
  saveExternalLinkValues?: boolean;
  /** Update links mode (CT_WorkbookPr `@updateLinks`) */
  updateLinks?: string;
  /** Show pivot chart filter (CT_WorkbookPr `@showPivotChartFilter`) */
  showPivotChartFilter?: boolean;
  /** Publish items (CT_WorkbookPr `@publishItems`) */
  publishItems?: boolean;
  /** Check compatibility (CT_WorkbookPr `@checkCompatibility`) */
  checkCompatibility?: boolean;
  /** Auto compress pictures (CT_WorkbookPr `@autoCompressPictures`) */
  autoCompressPictures?: boolean;
  /** Refresh all connections (CT_WorkbookPr `@refreshAllConnections`) */
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
}

/** Calculation properties (CT_CalcPr) */
export interface CalculationPropertiesOptions {
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
  /** Calc completed (CT_CalcPr `@calcCompleted`) */
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
  /** Workbook visibility: "visible" | "hidden" | "veryHidden" (default "visible") */
  visibility?: "visible" | "hidden" | "veryHidden";
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

/**
 * A named range, constant, or formula (CT_DefinedName, sml.xsd:4317).
 *
 * The element text holds the formula or reference (ST_Formula); the
 * attributes carry metadata (scope, visibility, function metadata).
 */
export interface DefinedNameOptions {
  /** Formula or reference (element text), e.g. "Sheet1!$A$1" or "SUM(Sheet1!A1:A10)" */
  value: string;
  /** Name (required) */
  name: string;
  /** Comment (CT_DefinedName `@comment`) */
  comment?: string;
  /** Custom menu text (CT_DefinedName `@customMenu`) */
  customMenu?: string;
  /** Description (CT_DefinedName `@description`) */
  description?: string;
  /** Help text (CT_DefinedName `@help`) */
  help?: string;
  /** Status bar text (CT_DefinedName `@statusBar`) */
  statusBar?: string;
  /** Local scope: 0-based sheet index; omit for workbook scope (CT_DefinedName `@localSheetId`) */
  localSheetId?: number;
  /** Hidden from the name manager (default false) */
  hidden?: boolean;
  /** User-defined function (default false) */
  function?: boolean;
  /** VBA procedure (default false) */
  vbProcedure?: boolean;
  /** XLM macro (default false) */
  xlm?: boolean;
  /** Function group ID (CT_DefinedName `@functionGroupId`) */
  functionGroupId?: number;
  /** Shortcut key (CT_DefinedName `@shortcutKey`) */
  shortcutKey?: string;
  /** Publish to SharePoint server (default false) */
  publishToServer?: boolean;
  /** Workbook parameter (default false) */
  workbookParameter?: boolean;
}

export interface FileVersionOptions {
  appName?: string;
  lastEdited?: number;
  lowestEdited?: number;
  rupBuild?: number;
}

/** ST_SmartTagShow — smart tag display policy (CT_SmartTagPr/`@show`). */
export type SmartTagShow = "all" | "none" | "noIndicator";

/** CT_SmartTagPr — workbook-level smart tag embed and display policy. */
export interface SmartTagPropertiesOptions {
  /** Embed smart tag data into the workbook (default false). */
  embed?: boolean;
  /** Smart tag display policy (default "all"). */
  show?: SmartTagShow;
}

/** CT_SmartTagType — a recognized smart tag namespace registered at workbook level. */
export interface SmartTagTypeOptions {
  /** Namespace URI identifying the smart tag recognizer. */
  namespaceUri?: string;
  /** Smart tag name within the namespace. */
  name?: string;
  /** Optional URL for more information about the smart tag. */
  url?: string;
}

export interface WorkbookDescriptorOptions {
  /** CT_FileVersion — Excel version stamp; fresh compiles emit Excel 2007 defaults. */
  fileVersion?: FileVersionOptions;
  sheets: SheetDefinition[];
  pivotCaches?: PivotCacheReference[];
  protection?: WorkbookProtectionOptions;
  customViews?: CustomWorkbookViewOptions[];
  /** CT_SmartTagPr — workbook-level smart tag embed/display policy. */
  smartTagPr?: SmartTagPropertiesOptions;
  /** CT_SmartTagTypes — recognized smart tag namespaces (smartTagType[]). */
  smartTagTypes?: SmartTagTypeOptions[];
  fileRecoveryPr?: FileRecoveryPropertiesOptions;
  functionGroups?: string[];
  webPublishing?: WebPublishingOptions;
  fileSharing?: FileSharingOptions;
  workbookPr?: WorkbookPropertiesOptions;
  calcPr?: CalculationPropertiesOptions;
  /** OLE embedded range (CT_OleSize, after calcPr per XSD sequence) */
  oleSize?: string;
  bookView?: WorkbookViewOptions;
  volTypes?: VolTypeOptions[];
  webPublishObjects?: WebPublishObjectOptions[];
  /** Defined names (CT_DefinedNames) — named ranges, constants, formulas */
  definedNames?: DefinedNameOptions[];
  conformance?: WorkbookConformance;
}
