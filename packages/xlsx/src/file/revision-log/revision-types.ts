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
  readonly id: number;
}

export interface RevisionHeaderEntry {
  /** Unique GUID for this revision */
  readonly guid: string;
  /** Date/time of revision (ISO 8601) */
  readonly dateTime: string;
  /** User who made the revision */
  readonly userName: string;
  /** Relationship ID pointing to the revision log file */
  readonly rId: string;
  /** Max sheet ID at time of revision */
  readonly maxSheetId: number;
  /** Sheet ID map entries */
  readonly sheetIds?: readonly SheetIdEntry[];
  /** Minimum revision ID */
  readonly minRId?: number;
  /** Maximum revision ID */
  readonly maxRId?: number;
}

export interface RevisionHeadersOptions {
  /** Unique GUID for the headers collection */
  readonly guid: string;
  /** Last GUID */
  readonly lastGuid?: string;
  /** Revision entries */
  readonly headers: readonly RevisionHeaderEntry[];
  /** Shared flag */
  readonly shared?: boolean;
  /** History flag */
  readonly history?: boolean;
  /** Track revisions flag */
  readonly trackRevisions?: boolean;
  /** Revision ID counter */
  readonly revisionId?: number;
  /** Version */
  readonly version?: number;
  /** Keep change history */
  readonly keepChangeHistory?: boolean;
  /** Protected */
  readonly protected?: boolean;
  /** Preserve history days */
  readonly preserveHistory?: number;
}

// ── Revision Log Options ──

/** Row/column insert/delete action */
export type RowColumnAction = "insertRow" | "insertCol" | "deleteRow" | "deleteCol";

export interface RevisionRowColumnOptions {
  /** Action type */
  readonly action: RowColumnAction;
  /** Revision ID */
  readonly rId: number;
  /** 1-based row number (for row actions) */
  readonly row?: number;
  /** 1-based column number (for column actions) */
  readonly col?: number;
  /** Sheet index (0-based) */
  readonly sheetIndex?: number;
  /** Edge (for insert edge tracking) */
  readonly edge?: boolean;
}

export interface RevisionCellChangeOptions {
  /** Revision ID */
  readonly rId: number;
  /** Cell reference (e.g., "A1") */
  readonly ref: string;
  /** Sheet index (0-based) */
  readonly sheetIndex?: number;
  /** Old value */
  readonly oldValue?: string | number;
  /** Old type ("b" | "d" | "e" | "n" | "s" | "str") */
  readonly oldType?: string;
  /** New value */
  readonly newValue?: string | number;
  /** New type */
  readonly newType?: string;
  /** Cell formula */
  readonly formula?: string;
  /** Number format id */
  readonly numFmtId?: number;
  /** XF differential format (CT_RevisionCellChange @xfDxf) */
  readonly xfDxf?: number;
  /** Quote prefix (CT_RevisionCellChange @quotePrefix) */
  readonly quotePrefix?: boolean;
  /** Old quote prefix */
  readonly oldQuotePrefix?: boolean;
  /** Phonetic (CT_RevisionCellChange @ph) */
  readonly ph?: boolean;
  /** Old phonetic */
  readonly oldPh?: boolean;
  /** End of list formula update */
  readonly endOfListFormulaUpdate?: boolean;
}

export interface RevisionMoveOptions {
  /** Revision ID */
  readonly rId: number;
  /** Source range (e.g., "A1:B2") */
  readonly source: string;
  /** Destination range (e.g., "C1:D2") */
  readonly destination: string;
  /** Sheet index (0-based) */
  readonly sheetIndex?: number;
  /** Source sheet ID (CT_RevisionMove @sourceSheetId) */
  readonly sourceSheetId?: number;
}

export interface RevisionFormattingOptions {
  /** Revision ID */
  readonly rId: number;
  /** Cell range (e.g., "A1:B2") */
  readonly ref: string;
  /** Sheet index (0-based) */
  readonly sheetIndex?: number;
  /** Style index */
  readonly s?: number;
  /** XF differential format (CT_RevisionFormatting @xfDxf) */
  readonly xfDxf?: number;
}

export interface RevisionInsertSheetOptions {
  /** Revision ID */
  readonly rId: number;
  /** Sheet name */
  readonly name: string;
  /** Sheet index */
  readonly sheetIndex?: number;
  /** Sheet position (CT_RevisionInsertSheet @sheetPosition) */
  readonly sheetPosition?: number;
}

export interface RevisionCommentOptions {
  /** Revision ID */
  readonly rId: number;
  /** Cell reference */
  readonly ref: string;
  /** Sheet index (0-based) */
  readonly sheetIndex?: number;
  /** Comment text */
  readonly text?: string;
  /** Author */
  readonly author?: string;
  /** Always show comment (CT_RevisionComment @alwaysShow) */
  readonly alwaysShow?: boolean;
  /** Old comment flag */
  readonly old?: boolean;
  /** Hidden row */
  readonly hiddenRow?: boolean;
  /** Hidden column */
  readonly hiddenColumn?: boolean;
  /** Old text length */
  readonly oldLength?: number;
  /** New text length */
  readonly newLength?: number;
}

export interface RevisionDefinedNameOptions {
  /** Revision ID */
  readonly rId: number;
  /** Defined name */
  readonly name: string;
  /** Formula/value */
  readonly value?: string;
  /** Local sheet ID (-1 for global) */
  readonly localSheetId?: number;
  /** Custom view (CT_RevisionDefinedName @customView) */
  readonly customView?: boolean;
  /** Is function (CT_RevisionDefinedName @function) */
  readonly function?: boolean;
  /** Old function flag */
  readonly oldFunction?: boolean;
  /** Function group ID */
  readonly functionGroupId?: number;
  /** Old function group ID */
  readonly oldFunctionGroupId?: number;
  /** Shortcut key */
  readonly shortcutKey?: string;
  /** Old shortcut key */
  readonly oldShortcutKey?: string;
  /** Old hidden flag */
  readonly oldHidden?: boolean;
  /** Custom menu */
  readonly customMenu?: string;
  /** Old custom menu */
  readonly oldCustomMenu?: string;
  /** Old description */
  readonly oldDescription?: string;
  /** Help text */
  readonly help?: string;
  /** Old help text */
  readonly oldHelp?: string;
  /** Status bar text */
  readonly statusBar?: string;
  /** Old status bar text */
  readonly oldStatusBar?: string;
  /** Old comment */
  readonly oldComment?: string;
}

/** Union type for all revision entries */
export type RevisionEntry =
  | { readonly type: "rowColumn"; readonly data: RevisionRowColumnOptions }
  | { readonly type: "cellChange"; readonly data: RevisionCellChangeOptions }
  | { readonly type: "move"; readonly data: RevisionMoveOptions }
  | { readonly type: "formatting"; readonly data: RevisionFormattingOptions }
  | { readonly type: "insertSheet"; readonly data: RevisionInsertSheetOptions }
  | { readonly type: "comment"; readonly data: RevisionCommentOptions }
  | { readonly type: "definedName"; readonly data: RevisionDefinedNameOptions };

export interface RevisionLogOptions {
  /** Revision entries */
  readonly revisions?: readonly RevisionEntry[];
}
