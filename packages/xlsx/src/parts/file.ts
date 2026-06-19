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
  FileRecoveryPropertiesOptions,
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
  /** Extended properties (docProps/app.xml) */
  appProperties?: AppPropertiesOptions;
  /** Custom properties (docProps/custom.xml); omitted from the package when empty */
  customProperties?: CustomPropertyOptions[];
}
