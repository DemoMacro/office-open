/**
 * WorkbookOptions type — the top-level options for XLSX generation.
 *
 * @module
 */

import type { CorePropertiesOptions } from "@parts/core-properties";
import type { DxfOptions } from "@parts/styles";
import type {
  WorkbookProtectionOptions,
  FileRecoveryPrOptions,
  WebPublishingOptions,
  FileSharingOptions,
  CustomWorkbookViewOptions,
  VolTypeOptions,
  WebPublishObjectOptions,
} from "@parts/workbook";

import type { ChartsheetOptions } from "./chartsheet";
import type { ExternalLinkOptions } from "./external-link";
import type { WorksheetOptions } from "./worksheet";

export interface WorkbookOptions extends CorePropertiesOptions {
  worksheets?: WorksheetOptions[];
  /** Chart-only sheets (no cells, just a chart) */
  chartsheets?: ChartsheetOptions[];
  /** Pre-defined differential formats for conditional formatting */
  dxfs?: DxfOptions[];
  /** Workbook-level protection */
  workbookProtection?: WorkbookProtectionOptions;
  /** External link definitions */
  externalLinks?: ExternalLinkOptions[];
  /** Custom workbook views */
  customWorkbookViews?: CustomWorkbookViewOptions[];
  /** File recovery properties */
  fileRecoveryPr?: FileRecoveryPrOptions;
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
}
