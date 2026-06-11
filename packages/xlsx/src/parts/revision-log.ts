/**
 * Revision types — shared interfaces for revision log XML generators.
 *
 * Reference: OOXML transitional, sml.xsd, CT_RevisionHeaders / CT_Revisions
 *
 * @module
 */

// ── Revision Headers Options ──

export interface SheetIdEntry {
  /** Sheet ID */
  id: number;
}

export interface RevisionHeaderEntry {
  /** Unique GUID for this revision */
  guid: string;
  /** Date/time of revision (ISO 8601) */
  dateTime: string;
  /** User who made the revision */
  userName: string;
  /** Relationship ID pointing to the revision log file */
  rId: string;
  /** Max sheet ID at time of revision */
  maxSheetId: number;
  /** Sheet ID map entries */
  sheetIds?: SheetIdEntry[];
  /** Minimum revision ID */
  minRId?: number;
  /** Maximum revision ID */
  maxRId?: number;
}

export interface RevisionHeadersOptions {
  /** Unique GUID for the headers collection */
  guid: string;
  /** Last GUID */
  lastGuid?: string;
  /** Revision entries */
  headers: RevisionHeaderEntry[];
  /** Shared flag */
  shared?: boolean;
  /** History flag */
  history?: boolean;
  /** Track revisions flag */
  trackRevisions?: boolean;
  /** Revision ID counter */
  revisionId?: number;
  /** Version */
  version?: number;
  /** Keep change history */
  keepChangeHistory?: boolean;
  /** Protected */
  protected?: boolean;
  /** Preserve history days */
  preserveHistory?: number;
  /** Disk revisions (CT_Headers @diskRevisions) */
  diskRevisions?: boolean;
  /** Exclusive (CT_Headers @exclusive) */
  exclusive?: boolean;
}

// ── Revision Log Options ──

/** Row/column insert/delete action */
export type RowColumnAction = "insertRow" | "insertCol" | "deleteRow" | "deleteCol";

export interface RevisionRowColumnOptions {
  /** Action type */
  action: RowColumnAction;
  /** Revision ID */
  rId: number;
  /** 1-based row number (for row actions) */
  row?: number;
  /** 1-based column number (for column actions) */
  col?: number;
  /** Sheet index (0-based) */
  sheetIndex?: number;
  /** Edge (for insert edge tracking) */
  edge?: boolean;
  /** Rejected (CT_RevisionRowColumn @ra) */
  ra?: string;
  /** Undo action (CT_RevisionRowColumn @ua) */
  ua?: string;
  /** End of list (CT_RevisionRowColumn @eol) */
  eol?: boolean;
}

export interface RevisionCellChangeOptions {
  /** Revision ID */
  rId: number;
  /** Cell reference (e.g., "A1") */
  ref: string;
  /** Sheet index (0-based) */
  sheetIndex?: number;
  /** Old value */
  oldValue?: string | number;
  /** Old type ("b" | "d" | "e" | "n" | "s" | "str") */
  oldType?: string;
  /** New value */
  newValue?: string | number;
  /** New type */
  newType?: string;
  /** Cell formula */
  formula?: string;
  /** Number format id */
  numFmtId?: number;
  /** XF differential format (CT_RevisionCellChange @xfDxf) */
  xfDxf?: number;
  /** Quote prefix (CT_RevisionCellChange @quotePrefix) */
  quotePrefix?: boolean;
  /** Old quote prefix */
  oldQuotePrefix?: boolean;
  /** Phonetic (CT_RevisionCellChange @ph) */
  ph?: boolean;
  /** Old phonetic */
  oldPh?: boolean;
  /** End of list formula update */
  endOfListFormulaUpdate?: boolean;
  /** Rejected (CT_RevisionCellChange @ra) */
  ra?: string;
  /** Undo action (CT_RevisionCellChange @ua) */
  ua?: string;
  /** Cell metadata for new cell (nc @cm) */
  newCellMeta?: number;
  /** Cell metadata for old cell (oc @cm) */
  oldCellMeta?: number;
}

export interface RevisionMoveOptions {
  /** Revision ID */
  rId: number;
  /** Source range (e.g., "A1:B2") */
  source: string;
  /** Destination range (e.g., "C1:D2") */
  destination: string;
  /** Sheet index (0-based) */
  sheetIndex?: number;
  /** Source sheet ID (CT_RevisionMove @sourceSheetId) */
  sourceSheetId?: number;
  /** Rejected (CT_RevisionMove @ra) */
  ra?: string;
  /** Undo action (CT_RevisionMove @ua) */
  ua?: string;
}

export interface RevisionFormattingOptions {
  /** Revision ID */
  rId: number;
  /** Cell range (e.g., "A1:B2") */
  ref: string;
  /** Sheet index (0-based) */
  sheetIndex?: number;
  /** Style index */
  s?: number;
  /** XF differential format (CT_RevisionFormatting @xfDxf) */
  xfDxf?: number;
}

export interface RevisionInsertSheetOptions {
  /** Revision ID */
  rId: number;
  /** Sheet name */
  name: string;
  /** Sheet index */
  sheetIndex?: number;
  /** Sheet position (CT_RevisionInsertSheet @sheetPosition) */
  sheetPosition?: number;
  /** Rejected (CT_RevisionInsertSheet @ra) */
  ra?: string;
  /** Undo action (CT_RevisionInsertSheet @ua) */
  ua?: string;
}

export interface RevisionCommentOptions {
  /** Revision ID */
  rId: number;
  /** Cell reference */
  ref: string;
  /** Sheet index (0-based) */
  sheetIndex?: number;
  /** Comment text */
  text?: string;
  /** Author */
  author?: string;
  /** Always show comment (CT_RevisionComment @alwaysShow) */
  alwaysShow?: boolean;
  /** Old comment flag */
  old?: boolean;
  /** Hidden row */
  hiddenRow?: boolean;
  /** Hidden column */
  hiddenColumn?: boolean;
  /** Old text length */
  oldLength?: number;
  /** New text length */
  newLength?: number;
}

export interface RevisionDefinedNameOptions {
  /** Revision ID */
  rId: number;
  /** Defined name */
  name: string;
  /** Formula/value */
  value?: string;
  /** Local sheet ID (-1 for global) */
  localSheetId?: number;
  /** Custom view (CT_RevisionDefinedName @customView) */
  customView?: boolean;
  /** Is function (CT_RevisionDefinedName @function) */
  function?: boolean;
  /** Old function flag */
  oldFunction?: boolean;
  /** Function group ID */
  functionGroupId?: number;
  /** Old function group ID */
  oldFunctionGroupId?: number;
  /** Shortcut key */
  shortcutKey?: string;
  /** Old shortcut key */
  oldShortcutKey?: string;
  /** Old hidden flag */
  oldHidden?: boolean;
  /** Custom menu */
  customMenu?: string;
  /** Old custom menu */
  oldCustomMenu?: string;
  /** Old description */
  oldDescription?: string;
  /** Help text */
  help?: string;
  /** Old help text */
  oldHelp?: string;
  /** Status bar text */
  statusBar?: string;
  /** Old status bar text */
  oldStatusBar?: string;
  /** Old comment */
  oldComment?: string;
  /** Rejected (CT_RevisionDefinedName @ra) */
  ra?: string;
  /** Undo action (CT_RevisionDefinedName @ua) */
  ua?: string;
}

/** Reviewed revision (CT_Reviewed) */
export interface RevisionReviewedOptions {
  /** Revision ID */
  rId: number;
}

/** Undo revision (CT_Undo) — wraps inner revision entries that were undone */
export interface RevisionUndoOptions {
  /** Revision ID of the undo action */
  rId: number;
  /** Revision entries wrapped by undo */
  revisions?: RevisionEntry[];
  /** Cell style (CT_Undo @cs) */
  cs?: boolean;
  /** Defined name (CT_Undo @dn) */
  dn?: boolean;
  /** Expansion (CT_Undo @exp) */
  exp?: boolean;
  /** Number format (CT_Undo @nf) */
  nf?: boolean;
  /** 3D reference (CT_Undo @ref3D) */
  ref3D?: boolean;
}

/** Auto formatting revision (CT_RevisionAutoFormatting) */
export interface RevisionAutoFormattingOptions {
  /** Revision ID */
  rId: number;
  /** Sheet index */
  sheetIndex?: number;
  /** Cell reference */
  ref?: string;
}

/** Custom view revision (CT_RevisionCustomView) */
export interface RevisionCustomViewOptions {
  /** Revision ID */
  rId: number;
  /** GUID */
  guid?: string;
}

/** Sheet rename revision (CT_RevisionSheetRename) */
export interface RevisionSheetRenameOptions {
  /** Revision ID */
  rId: number;
  /** Sheet index */
  sheetIndex?: number;
  /** Old sheet name */
  oldName?: string;
  /** New sheet name */
  newName?: string;
  /** Rejected (CT_RevisionSheetRename @ra) */
  ra?: string;
  /** Undo action (CT_RevisionSheetRename @ua) */
  ua?: string;
}

/** Query table field revision (CT_RevisionQueryTableField) */
export interface RevisionQueryTableFieldOptions {
  /** Revision ID */
  rId: number;
  /** Sheet index */
  sheetIndex?: number;
  /** Field ID */
  fieldId?: number;
  /** Rejected (CT_RevisionQueryTableField @ra) */
  ra?: string;
  /** Undo action (CT_RevisionQueryTableField @ua) */
  ua?: string;
}

/** Conflict revision (CT_RevisionConflict) */
export interface RevisionConflictOptions {
  /** Revision ID */
  rId: number;
  /** Sheet index */
  sheetIndex?: number;
  /** Rejected (CT_RevisionConflict @ra) */
  ra?: string;
  /** Undo action (CT_RevisionConflict @ua) */
  ua?: string;
}

/** Union type for all revision entries */
export type RevisionEntry =
  | { type: "rowColumn"; data: RevisionRowColumnOptions }
  | { type: "cellChange"; data: RevisionCellChangeOptions }
  | { type: "move"; data: RevisionMoveOptions }
  | { type: "formatting"; data: RevisionFormattingOptions }
  | { type: "insertSheet"; data: RevisionInsertSheetOptions }
  | { type: "comment"; data: RevisionCommentOptions }
  | { type: "definedName"; data: RevisionDefinedNameOptions }
  | { type: "reviewed"; data: RevisionReviewedOptions }
  | { type: "undo"; data: RevisionUndoOptions }
  | { type: "autoFormatting"; data: RevisionAutoFormattingOptions }
  | { type: "customView"; data: RevisionCustomViewOptions }
  | { type: "sheetRename"; data: RevisionSheetRenameOptions }
  | { type: "queryTableField"; data: RevisionQueryTableFieldOptions }
  | { type: "conflict"; data: RevisionConflictOptions };

export interface RevisionLogOptions {
  /** Revision entries */
  revisions?: RevisionEntry[];
}

/** Shared user info (CT_SharedUser) */
export interface SharedUserOptions {
  /** GUID (required) */
  guid: string;
  /** User name (required) */
  name: string;
  /** User ID (required) */
  id: number;
  /** Date of last edit */
  dateTime?: string;
}

/** Users collection (CT_Users) */
export interface UsersOptions {
  /** User entries */
  users?: SharedUserOptions[];
}
