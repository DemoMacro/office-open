/**
 * Dialogsheet options for SpreadsheetML documents.
 *
 * A dialogsheet is a legacy Excel 5.0 dialog sheet (no cell data).
 *
 * Reference: OOXML transitional, sml.xsd, CT_Dialogsheet
 *
 * @module
 */

export interface DialogsheetPageMargins {
  left?: number;
  right?: number;
  top?: number;
  bottom?: number;
  header?: number;
  footer?: number;
}

export interface DialogsheetPageSetup {
  paperSize?: number;
  orientation?: string;
  horizontalDpi?: number;
  verticalDpi?: number;
  copies?: number;
}

export interface DialogsheetProtectionOptions {
  content?: boolean;
  objects?: boolean;
  scenarios?: boolean;
}

export interface DialogsheetOptions {
  /** Sheet name */
  name?: string;
  /** Tab color (hex ARGB) */
  tabColor?: string;
  /** Page margins */
  pageMargins?: DialogsheetPageMargins;
  /** Page setup */
  pageSetup?: DialogsheetPageSetup;
  /** Sheet protection */
  sheetProtection?: DialogsheetProtectionOptions;
}
