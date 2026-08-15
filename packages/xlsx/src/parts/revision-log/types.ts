/**
 * Revision log — option types for shared-workbook change tracking.
 *
 * @module
 */

// ── Revision row/column action (ST_rwColActionType, sml.xsd:2076) ──

export type RowColumnAction = "insertRow" | "deleteRow" | "insertCol" | "deleteCol";

/** Revision action for comments / custom views (ST_RevisionAction, sml.xsd:2084). */
export type RevisionAction = "add" | "delete";

// ── Revision Headers — xl/revisionHeaders.xml (CT_RevisionHeaders) ──

/** One header entry (CT_RevisionHeader, sml.xsd:1898). */
export interface RevisionHeaderEntry {
  /** Unique GUID for this revision (required, ST_Guid). */
  guid: string;
  /** Date/time of the revision (required, xsd:dateTime). */
  dateTime: string;
  /** User who made the revision (required, ST_Xstring). */
  userName: string;
  /** Relationship ID pointing to the revision log file (required, r:id). */
  rId: string;
  /** Max sheet ID at the time of revision (required). */
  maxSheetId: number;
  /** Sheet ID map entries (required, ≥1). */
  sheetIds: number[];
  /** Reviewed revision IDs (reviewedList.reviewed[].rId, optional). */
  reviewed?: number[];
  /** Minimum revision ID (optional). */
  minRId?: number;
  /** Maximum revision ID (optional). */
  maxRId?: number;
}

/** Options for xl/revisionHeaders.xml (CT_RevisionHeaders, sml.xsd:1860). */
export interface RevisionHeadersOptions {
  /** Unique GUID for the headers collection (required). */
  guid: string;
  /** Revision header entries (required, ≥1). */
  headers: RevisionHeaderEntry[];
  /** Last GUID (optional). */
  lastGuid?: string;
  /** Shared workbook (default true). */
  shared?: boolean;
  /** Disk revisions (default false). */
  diskRevisions?: boolean;
  /** History kept (default true). */
  history?: boolean;
  /** Track revisions (default true). */
  trackRevisions?: boolean;
  /** Exclusive (default false). */
  exclusive?: boolean;
  /** Revision ID counter (default 0). */
  revisionId?: number;
  /** Version (default 1). */
  version?: number;
  /** Keep change history (default true). */
  keepChangeHistory?: boolean;
  /** Protected (default false). */
  protected?: boolean;
  /** Preserve history days (default 30). */
  preserveHistory?: number;
}

// ── Users — xl/users.xml (CT_Users, sml.xsd:2100) ──

/** One shared-workbook user (CT_SharedUser, sml.xsd:2106). */
export interface SharedUserOptions {
  /** GUID (required, ST_Guid). */
  guid: string;
  /** User name (required, ST_Xstring). */
  name: string;
  /** User ID (required). */
  id: number;
  /** Date/time of last edit (required, xsd:dateTime). */
  dateTime: string;
}

/** Options for xl/users.xml (CT_Users, sml.xsd:2100). */
export interface UsersOptions {
  /** Shared-user entries (optional, ≤256). */
  users?: SharedUserOptions[];
}

// ── Revision Log — xl/revisions/revisionN.xml (CT_Revisions, sml.xsd:1877) ──

// AG_RevData (sml.xsd:1893): rId required, ua (undo), ra (rejected).
// Present on: rrc, rm, rsnm, ris, rcc, rdn, rcft. Absent on: rcv, rfmt, raf, rcmt, rqt.

/** Undo tracking info (CT_UndoInfo, sml.xsd:1933). */
export interface RevisionUndoOptions {
  /** Undo index (required). */
  index: number;
  /** Expression kind (required, ST_FormulaExpression). */
  expression: "ref" | "refError" | "area" | "areaError" | "computedArea";
  /** Undo range (required, ST_RefA). */
  dr: string;
  /** 3D reference (default false). */
  ref3D?: boolean;
  /** Array formula (default false). */
  array?: boolean;
  /** Value (default false). */
  v?: boolean;
  /** Number format (default false). */
  nf?: boolean;
  /** Conditional style (default false). */
  cs?: boolean;
  /** Defined name (optional). */
  dn?: string;
  /** Cell reference (optional). */
  r?: string;
  /** Sheet ID (optional). */
  sId?: number;
}

/** Nested rrc/rm child — undo tracking, cell change, or formatting (shared choice group). */
export type RevisionNestedChild =
  | { kind: "undo"; data: RevisionUndoOptions }
  | { kind: "cellChange"; data: RevisionCellChangeOptions }
  | { kind: "formatting"; data: RevisionFormattingOptions };

/** Row/column insert/delete (CT_RevisionRowColumn, sml.xsd:1943). */
export interface RevisionRowColumnOptions {
  /** Revision ID (AG_RevData rId, required). */
  rId: number;
  /** Sheet ID (sId, required). */
  sheetId: number;
  /** Affected range (required, ST_Ref). */
  ref: string;
  /** Insertion/deletion action (required). */
  action: RowColumnAction;
  /** End of list (eol, default false). */
  endOfList?: boolean;
  /** Edge tracking (default false). */
  edge?: boolean;
  /** Revision was performed by an undo (ua). */
  undo?: boolean;
  /** Revision was rejected (ra). */
  rejected?: boolean;
  /** Nested undo/rcc/rfmt children (optional). */
  children?: RevisionNestedChild[];
}

/** Cell move (CT_RevisionMove, sml.xsd:1956). */
export interface RevisionMoveOptions {
  /** Revision ID (AG_RevData rId, required). */
  rId: number;
  /** Sheet ID (required). */
  sheetId: number;
  /** Source range (required, ST_Ref). */
  source: string;
  /** Destination range (required, ST_Ref). */
  destination: string;
  /** Source sheet ID (default 0). */
  sourceSheetId?: number;
  /** Revision was performed by an undo (ua). */
  undo?: boolean;
  /** Revision was rejected (ra). */
  rejected?: boolean;
  /** Nested undo/rcc/rfmt children (optional). */
  children?: RevisionNestedChild[];
}

/** Custom view add/delete (CT_RevisionCustomView, sml.xsd:1968 — no AG_RevData). */
export interface RevisionCustomViewOptions {
  /** GUID (required, ST_Guid). */
  guid: string;
  /** Action (required). */
  action: RevisionAction;
}

/** Sheet rename (CT_RevisionSheetRename, sml.xsd:1972). */
export interface RevisionSheetRenameOptions {
  /** Revision ID (AG_RevData rId, required). */
  rId: number;
  /** Sheet ID (required). */
  sheetId: number;
  /** Old sheet name (required, ST_Xstring). */
  oldName: string;
  /** New sheet name (required, ST_Xstring). */
  newName: string;
  /** Revision was performed by an undo (ua). */
  undo?: boolean;
  /** Revision was rejected (ra). */
  rejected?: boolean;
}

/** Sheet insertion (CT_RevisionInsertSheet, sml.xsd:1981). */
export interface RevisionInsertSheetOptions {
  /** Revision ID (AG_RevData rId, required). */
  rId: number;
  /** Sheet ID (required). */
  sheetId: number;
  /** Sheet name (required, ST_Xstring). */
  name: string;
  /** Sheet position (required). */
  sheetPosition: number;
  /** Revision was performed by an undo (ua). */
  undo?: boolean;
  /** Revision was rejected (ra). */
  rejected?: boolean;
}

/** Cell value change (CT_RevisionCellChange, sml.xsd:1987). */
export interface RevisionCellChangeOptions {
  /** Revision ID (AG_RevData rId, required). */
  rId: number;
  /** Sheet ID (sId, required). */
  sheetId: number;
  /** Has old differential format (odxf attr, default false). */
  hasOldDxf?: boolean;
  /** XF differential format (xfDxf, default false). */
  xfDxf?: boolean;
  /** Has style (s, default false). */
  style?: boolean;
  /** Has differential format (dxf attr, default false). */
  hasDxf?: boolean;
  /** Number format ID (optional). */
  numFmtId?: number;
  /** Quote prefix (default false). */
  quotePrefix?: boolean;
  /** Old quote prefix (default false). */
  oldQuotePrefix?: boolean;
  /** Phonetic (ph, default false). */
  phonetic?: boolean;
  /** Old phonetic (default false). */
  oldPhonetic?: boolean;
  /** End of list formula update (default false). */
  endOfListFormulaUpdate?: boolean;
  /** Revision was performed by an undo (ua). */
  undo?: boolean;
  /** Revision was rejected (ra). */
  rejected?: boolean;
  /** Old cell (oc) verbatim CT_Cell XML (optional). */
  oldCellXml?: string;
  /** New cell (nc) verbatim CT_Cell XML (required). */
  newCellXml: string;
  /** Old differential format (odxf) verbatim CT_Dxf XML (optional). */
  oldDxfXml?: string;
  /** New differential format (ndxf) verbatim CT_Dxf XML (optional). */
  newDxfXml?: string;
}

/** Formatting change (CT_RevisionFormatting, sml.xsd:2008 — no AG_RevData). */
export interface RevisionFormattingOptions {
  /** Sheet ID (required). */
  sheetId: number;
  /** Sequence of cell refs (sqref, required). */
  sqref: string;
  /** XF differential format (default false). */
  xfDxf?: boolean;
  /** Has style (s, default false). */
  style?: boolean;
  /** Start index (optional). */
  start?: number;
  /** Length (optional). */
  length?: number;
  /** Differential format (dxf) verbatim CT_Dxf XML (optional). */
  dxfXml?: string;
}

/** Auto formatting (CT_RevisionAutoFormatting, sml.xsd:2020 — no AG_RevData). */
export interface RevisionAutoFormattingOptions {
  /** Sheet ID (required). */
  sheetId: number;
  /** Affected range (required, ST_Ref). */
  ref: string;
  /** AG_AutoFormat attributes verbatim (optional, e.g. format/align/...). */
  autoFormatXml?: string;
}

/** Defined name change (CT_RevisionDefinedName, sml.xsd:2038). */
export interface RevisionDefinedNameOptions {
  /** Revision ID (AG_RevData rId, required). */
  rId: number;
  /** Defined name (required, ST_Xstring). */
  name: string;
  /** Local sheet ID (optional; -1 or omitted = workbook-global). */
  localSheetId?: number;
  /** Custom view (default false). */
  customView?: boolean;
  /** New formula expression (optional). */
  formula?: string;
  /** Old formula expression (optional). */
  oldFormula?: string;
  /** Is function (default false). */
  function?: boolean;
  /** Old function flag (default false). */
  oldFunction?: boolean;
  /** Function group ID (optional). */
  functionGroupId?: number;
  /** Old function group ID (optional). */
  oldFunctionGroupId?: number;
  /** Shortcut key (optional, unsignedByte). */
  shortcutKey?: number;
  /** Old shortcut key (optional, unsignedByte). */
  oldShortcutKey?: number;
  /** Hidden (default false). */
  hidden?: boolean;
  /** Old hidden flag (default false). */
  oldHidden?: boolean;
  /** Custom menu (optional). */
  customMenu?: string;
  /** Old custom menu (optional). */
  oldCustomMenu?: string;
  /** Description (optional). */
  description?: string;
  /** Old description (optional). */
  oldDescription?: string;
  /** Help text (optional). */
  help?: string;
  /** Old help text (optional). */
  oldHelp?: string;
  /** Status bar text (optional). */
  statusBar?: string;
  /** Old status bar text (optional). */
  oldStatusBar?: string;
  /** Comment (optional). */
  comment?: string;
  /** Old comment (optional). */
  oldComment?: string;
  /** Revision was performed by an undo (ua). */
  undo?: boolean;
  /** Revision was rejected (ra). */
  rejected?: boolean;
}

/** Comment revision (CT_RevisionComment, sml.xsd:2025 — no AG_RevData). */
export interface RevisionCommentOptions {
  /** Sheet ID (required). */
  sheetId: number;
  /** Cell reference (required, ST_CellRef). */
  cell: string;
  /** GUID (required, ST_Guid). */
  guid: string;
  /** Action (default add). */
  action?: RevisionAction;
  /** Always show (default false). */
  alwaysShow?: boolean;
  /** Old comment flag (default false). */
  old?: boolean;
  /** Hidden row (default false). */
  hiddenRow?: boolean;
  /** Hidden column (default false). */
  hiddenColumn?: boolean;
  /** Author (required, ST_Xstring). */
  author: string;
  /** Old text length (default 0). */
  oldLength?: number;
  /** New text length (default 0). */
  newLength?: number;
}

/** Query table field change (CT_RevisionQueryTableField, sml.xsd:2071 — no AG_RevData). */
export interface RevisionQueryTableFieldOptions {
  /** Sheet ID (required). */
  sheetId: number;
  /** Affected range (required, ST_Ref). */
  ref: string;
  /** Field ID (required). */
  fieldId: number;
}

/** Conflict (CT_RevisionConflict, sml.xsd:2067). */
export interface RevisionConflictOptions {
  /** Revision ID (AG_RevData rId, required). */
  rId: number;
  /** Revision was performed by an undo (ua). */
  undo?: boolean;
  /** Revision was rejected (ra). */
  rejected?: boolean;
  /** Sheet ID (optional). */
  sheetId?: number;
}

/** Discriminated union of all CT_Revisions choice elements (sml.xsd:1877-1891). */
export type RevisionEntry =
  | { type: "rowColumn"; data: RevisionRowColumnOptions }
  | { type: "move"; data: RevisionMoveOptions }
  | { type: "customView"; data: RevisionCustomViewOptions }
  | { type: "sheetRename"; data: RevisionSheetRenameOptions }
  | { type: "insertSheet"; data: RevisionInsertSheetOptions }
  | { type: "cellChange"; data: RevisionCellChangeOptions }
  | { type: "formatting"; data: RevisionFormattingOptions }
  | { type: "autoFormatting"; data: RevisionAutoFormattingOptions }
  | { type: "definedName"; data: RevisionDefinedNameOptions }
  | { type: "comment"; data: RevisionCommentOptions }
  | { type: "queryTableField"; data: RevisionQueryTableFieldOptions }
  | { type: "conflict"; data: RevisionConflictOptions };

/** Options for xl/revisions/revisionN.xml (CT_Revisions, sml.xsd:1877). */
export interface RevisionLogOptions {
  /** Revision entries (the CT_Revisions choice sequence). */
  revisions: RevisionEntry[];
}
