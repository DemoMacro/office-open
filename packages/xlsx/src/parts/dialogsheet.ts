/**
 * Dialogsheet options for SpreadsheetML documents.
 *
 * A dialogsheet is a legacy Excel 5.0 dialog sheet (no cell data).
 *
 * Reference: OOXML transitional, sml.xsd, CT_Dialogsheet
 *
 * @module
 */

import type { UniversalMeasure } from "@office-open/core";

export interface DialogsheetPageMargins {
  left?: number | UniversalMeasure;
  right?: number | UniversalMeasure;
  top?: number | UniversalMeasure;
  bottom?: number | UniversalMeasure;
  header?: number | UniversalMeasure;
  footer?: number | UniversalMeasure;
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
