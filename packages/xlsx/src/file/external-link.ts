/**
 * External Link options for SpreadsheetML documents.
 *
 * Reference: OOXML transitional, sml.xsd, CT_ExternalLink
 *
 * @module
 */

// ── Options ──

export interface ExternalDefinedNameOptions {
  name: string;
  refersTo?: string;
  sheetId?: number;
  /** Publish to server (CT_DefinedName @publishToServer) */
  publishToServer?: boolean;
  /** VBA procedure (CT_DefinedName @vbProcedure) */
  vbProcedure?: boolean;
  /** Workbook parameter (CT_DefinedName @workbookParameter) */
  workbookParameter?: boolean;
  /** XLM macro (CT_DefinedName @xlm) */
  xlm?: boolean;
}

export interface ExternalCellOptions {
  /** Cell reference, e.g. "A1" */
  reference: string;
  /** Cell data type */
  type?: string;
  /** Cell value */
  value?: string;
}

export interface ExternalBookOptions {
  /** Target path of the external workbook */
  target?: string;
  /** Sheet names from the external workbook */
  sheetNames?: string[];
  /** Defined names from the external workbook */
  definedNames?: ExternalDefinedNameOptions[];
  /** Cached sheet data from the external workbook */
  sheetDataSet?: ExternalSheetDataOptions[];
}

export interface ExternalRowOptions {
  /** Row number (1-based) */
  rowNumber: number;
  cells?: ExternalCellOptions[];
}

export interface ExternalSheetDataOptions {
  sheetId: number;
  refreshError?: boolean;
  rows?: ExternalRowOptions[];
}

export interface ExternalLinkOptions {
  /** External book configuration */
  externalBook?: ExternalBookOptions;
  /** Relationship ID for the external book (set by compiler) */
  bookRId?: string;
  /** OLE link configuration (CT_OleLink) */
  oleLink?: OleLinkOptions;
  /** Relationship ID for the OLE link (set by compiler) */
  oleRId?: string;
}

export interface OleItemOptions {
  /** OLE item name (required) */
  name: string;
  /** Whether to advise events */
  advise?: boolean;
  /** Whether preferred */
  prefer?: boolean;
}

export interface OleLinkOptions {
  /** OLE items */
  oleItems?: OleItemOptions[];
}
