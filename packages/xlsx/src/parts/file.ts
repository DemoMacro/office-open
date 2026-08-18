/**
 * WorkbookOptions type — the top-level options for XLSX generation.
 *
 * @module
 */

import type {
  AppPropertiesOptions,
  CorePropertiesOptions,
  CustomPropertyOptions,
  DataType,
} from "@office-open/core";
import type { ThemeOptions } from "@office-open/core/theme";
import type {
  CellStyleXfOptions,
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
  DefinedNameOptions,
} from "@parts/workbook";

import type { CalcCell } from "./calc-chain";
import type { ChartsheetOptions } from "./chartsheet";
import type { ConnectionOptions } from "./connection";
import type { DialogsheetOptions } from "./dialogsheet";
import type { ExternalLinkOptions } from "./external-link";
import type { MetadataOptions } from "./metadata";
import type { PivotCacheDefParseResult, PivotCacheRecordsParseResult } from "./pivot-cache";
import type { RevisionHeadersOptions, RevisionLogOptions, UsersOptions } from "./revision-log";
import type { RichTextOptions, WorksheetOptions } from "./worksheet";
import type { MapInfoOptions } from "./xml-mapping";

export interface WorkbookOptions extends CorePropertiesOptions {
  worksheets?: WorksheetOptions[];
  /** Chart-only sheets (no cells, just a chart) */
  chartsheets?: ChartsheetOptions[];
  /** Legacy Excel 5.0 dialog sheets (xl/dialogSheets/sheetN.xml) */
  dialogsheets?: DialogsheetOptions[];
  /** Pre-defined differential formats for conditional formatting */
  dxfs?: DxfOptions[];
  /** Custom color palette (CT_Colors) */
  colors?: ColorsOptions;
  /** Custom table/pivot table styles (CT_TableStyles) */
  tableStyles?: CustomTableStyleOptions[];
  /** Theme (xl/theme/theme1.xml) — structured round-trip; fresh default when omitted */
  theme?: ThemeOptions;
  /** Custom named cell styles (CT_CellStyles) */
  cellStyles?: CustomCellStyleOptions[];
  /** Named cell-style templates (CT_CellStyleXfs) — definitions referenced by cellStyles */
  cellStyleXfs?: CellStyleXfOptions[];
  /** Style sheet extensions (CT_ExtensionList on styleSheet) */
  styleExtensions?: StyleExtensionOptions[];
  /** Workbook-level protection */
  workbookProtection?: WorkbookProtectionOptions;
  /** External link definitions */
  externalLinks?: ExternalLinkOptions[];
  /** Workbook data connections (xl/connections.xml) */
  connections?: ConnectionOptions[];
  /** Rich metadata block (xl/metadata.xml) */
  metadata?: MetadataOptions;
  /** XML mappings (xl/xmlMaps.xml) */
  xmlMaps?: MapInfoOptions;
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
  /** Defined names — named ranges, constants, formulas (CT_DefinedNames) */
  definedNames?: DefinedNameOptions[];
  /** Workbook view (CT_BookView) */
  bookView?: WorkbookViewOptions;
  /** Calculation properties (CT_CalcPr) */
  calcPr?: CalculationPropertiesOptions;
  /** OLE embedded range (CT_OleSize) — workbook-level, after calcPr */
  oleSize?: string;
  /** Workbook properties (CT_WorkbookPr) — includes defaultThemeVersion, which Excel
   * requires when the package carries a theme (dropping it makes the file unopenable) */
  workbookPr?: WorkbookPropertiesOptions;
  /** Calculation chain cells (xl/calcChain.xml) — parse-only; compiler rebuilds from formulas */
  calcChain?: CalcCell[];
  /** Pivot cache definitions — parse-only (CT-layer); compiler regenerates from sourceData */
  pivotCaches?: PivotCacheDefParseResult[];
  /** Pivot cache records — parse-only (CT-layer) */
  pivotCacheRecords?: PivotCacheRecordsParseResult[];
  /**
   * Shared string table (xl/sharedStrings.xml), in table order. Parse fills
   * it so generate() can seed the write context from the source entries —
   * cells referencing rich-text si values keep their structure and index
   * instead of collapsing to flattened text. Fresh documents omit it and
   * register strings from cell values as they stringify.
   */
  sharedStrings?: (string | RichTextOptions)[];
  /** Extended properties (docProps/app.xml) */
  appProperties?: AppPropertiesOptions;
  /** Custom properties (docProps/custom.xml); omitted from the package when empty */
  customProperties?: CustomPropertyOptions[];
  /** Shared-workbook revision log (xl/revisionHeaders.xml + xl/revisions/revisionN.xml + xl/users.xml). */
  revisionLog?: SharedWorkbookOptions;
  /**
   * Parts carried verbatim from the source that generate() does not rebuild
   * (drawings, VML, external links, any unknown extension part). Collected
   * wholesale by the core passthrough pipeline — everything the model did not
   * absorb survives with bytes and content-type declaration intact. Parts the
   * compiler happens to rebuild under the same path win over the passthrough
   * copy (assembly order), so model-driven parts need no exclusion here.
   */
  rawParts?: { path: string; data: DataType; contentType?: string }[];
  /**
   * Relationships from rebuilt parts' source .rels that point at rawParts
   * (e.g. workbook.xml → externalLink/pivotCache). Re-emitted verbatim —
   * target unchanged (passthrough paths never move), fresh rId. Round-trip only.
   */
  passthroughRelationships?: { source: string; relationshipType: string; target: string }[];
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
