/**
 * WorkbookOptions type — the top-level options for XLSX generation.
 *
 * @module
 */

import type {
  AppPropertiesOptions,
  CorePropertiesOptions,
  CustomPropertyOptions,
} from "@office-open/core";
import type {
  ColorsOptions,
  CustomCellStyleOptions,
  CustomTableStyleOptions,
  DxfOptions,
  StyleExtensionOptions,
} from "@parts/styles";
import type {
  WorkbookProtectionOptions,
  WorkbookViewOptions,
  CalculationPropertiesOptions,
  WorkbookPropertiesOptions,
  FileRecoveryPropertiesOptions,
  WebPublishingOptions,
  FileSharingOptions,
  CustomWorkbookViewOptions,
  VolTypeOptions,
  WebPublishObjectOptions,
} from "@parts/workbook";

import type { CalcCell } from "./calc-chain";
import type { ChartsheetOptions } from "./chartsheet";
import type { ExternalLinkOptions } from "./external-link";
import type { PivotCacheDefParseResult, PivotCacheRecordsParseResult } from "./pivot-cache";
import type { RevisionHeadersOptions, RevisionLogOptions, UsersOptions } from "./revision-log";
import type { WorksheetOptions } from "./worksheet";

export interface WorkbookOptions extends CorePropertiesOptions {
  worksheets?: WorksheetOptions[];
  /** Chart-only sheets (no cells, just a chart) */
  chartsheets?: ChartsheetOptions[];
  /** Pre-defined differential formats for conditional formatting */
  dxfs?: DxfOptions[];
  /** Custom color palette (CT_Colors) */
  colors?: ColorsOptions;
  /** Custom table/pivot table styles (CT_TableStyles) */
  tableStyles?: CustomTableStyleOptions[];
  /** Custom named cell styles (CT_CellStyles) */
  cellStyles?: CustomCellStyleOptions[];
  /** Style sheet extensions (CT_ExtensionList on styleSheet) */
  styleExtensions?: StyleExtensionOptions[];
  /** Workbook-level protection */
  workbookProtection?: WorkbookProtectionOptions;
  /** External link definitions */
  externalLinks?: ExternalLinkOptions[];
  /** Custom workbook views */
  customWorkbookViews?: CustomWorkbookViewOptions[];
  /** File recovery properties */
  fileRecoveryPr?: FileRecoveryPropertiesOptions;
  /** Custom VBA function group names */
  functionGroups?: string[];
  /** Web publishing properties */
  webPublishing?: WebPublishingOptions;
  /** File sharing / read-only recommendation */
  fileSharing?: FileSharingOptions;
  /** Volatile dependencies (CT_VolTypes) */
  volTypes?: VolTypeOptions[];
  /** Web publish objects (CT_WebPublishItems) */
  webPublishObjects?: WebPublishObjectOptions[];
  /** Workbook view (CT_BookView) — parse-only; compiler does not round-trip this field */
  bookView?: WorkbookViewOptions;
  /** Calculation properties (CT_CalcPr) — parse-only */
  calcPr?: CalculationPropertiesOptions;
  /** Workbook properties (CT_WorkbookPr) — parse-only */
  workbookPr?: WorkbookPropertiesOptions;
  /** Calculation chain cells (xl/calcChain.xml) — parse-only; compiler rebuilds from formulas */
  calcChain?: CalcCell[];
  /** Pivot cache definitions — parse-only (CT-layer); compiler regenerates from sourceData */
  pivotCaches?: PivotCacheDefParseResult[];
  /** Pivot cache records — parse-only (CT-layer) */
  pivotCacheRecords?: PivotCacheRecordsParseResult[];
  /** Extended properties (docProps/app.xml) */
  appProperties?: AppPropertiesOptions;
  /** Custom properties (docProps/custom.xml); omitted from the package when empty */
  customProperties?: CustomPropertyOptions[];
  /** Shared-workbook revision log (xl/revisionHeaders.xml + xl/revisions/revisionN.xml + xl/users.xml). */
  revisionLog?: SharedWorkbookOptions;
}

/** Shared-workbook revision tracking data. */
export interface SharedWorkbookOptions {
  /** xl/revisionHeaders.xml (CT_RevisionHeaders). */
  headers: RevisionHeadersOptions;
  /** Revision logs (CT_Revisions), one per header entry. logs[i] maps to headers.headers[i].rId. */
  logs: RevisionLogOptions[];
  /** xl/users.xml (CT_Users), optional. */
  users?: UsersOptions;
}
