/**
 * WorkbookOptions type — the top-level options for XLSX generation.
 *
 * @module
 */

import type {
  AppPropertiesOptions,
  ContentTypesInput,
  CorePropertiesOptions,
  CustomPropertyOptions,
  DataType,
  EncryptedContainerOptions,
  FormContainerOptions,
} from "@office-open/core";
import type { ThemeOptions } from "@office-open/core/theme";
import type {
  CellFillOptions,
  CellStyleXfOptions,
  ColorsOptions,
  CustomCellStyleOptions,
  TableStylesInfo,
  DxfOptions,
  FontOptions,
  FontsContainerOptions,
  BorderSideOptions,
  IndexedXfEntry,
  NumFmtEntry,
  StyleExtensionOptions,
} from "@parts/styles";
import type {
  WorkbookProtectionOptions,
  WorkbookViewOptions,
  WorkbookConformance,
  CalculationPropertiesOptions,
  WorkbookPropertiesOptions,
  FileRecoveryPropertiesOptions,
  WebPublishingOptions,
  FileSharingOptions,
  CustomWorkbookViewOptions,
  FileVersionOptions,
  VolTypeOptions,
  WebPublishObjectOptions,
  DefinedNameOptions,
  SheetDefinition,
} from "@parts/workbook";

import type { CalcCell } from "./calc-chain";
import type { ChartsheetOptions } from "./chartsheet";
import type { ClassificationLabelListOptions } from "./classification-labels";
import type { ConnectionOptions } from "./connection";
import type { DialogsheetOptions } from "./dialogsheet";
import type { ExternalLinkOptions } from "./external-link";
import type { MetadataOptions } from "./metadata";
import type { PersonsOptions } from "./persons";
import type { PivotSourceData } from "./pivot";
import type { PivotCacheDefinitionOptions, PivotCacheRecordsOptions } from "./pivot-cache";
import type { RevisionHeadersOptions, RevisionLogOptions, UsersOptions } from "./revision-log";
import type { RichDataPackageOptions } from "./rich-data";
import type { RevisionPtrOptions, WorkbookExtensionOptions } from "./workbook";
import type { RichTextOptions, WorksheetOptions } from "./worksheet";
import type { MapInfoOptions } from "./xml-mapping";

/** Root options for a .xlsx package — maps to workbook.xml plus its satellite parts. */
export interface WorkbookOptions extends CorePropertiesOptions {
  /** Source primary workbook part path; default xl/workbook.xml. Round-trip only. */
  workbookPath?: string;
  /** Source styles part path; default xl/styles.xml. Round-trip only. */
  stylesPath?: string;
  /** Source sharedStrings part path; default xl/sharedStrings.xml. Round-trip only. */
  sharedStringsPath?: string;
  /** Source theme part path; default xl/theme/theme1.xml. Round-trip only. */
  themePath?: string;
  worksheets?: WorksheetOptions[];
  /** Workbook <sheets> order; round-trip only. */
  sheetDefinitions?: SheetDefinition[];
  /** Workbook-level form containers. */
  forms?: FormContainerOptions[];
  /**
   * Source file is an encrypted OOXML package (OLE2/CFB container). Round-trip
   * only: the original bytes are re-emitted verbatim; all other fields stay
   * empty — mixing sheet content is rejected.
   */
  encrypted?: EncryptedContainerOptions;
  /**
   * Workbook folder path (x15ac:absPath/@url in an mc:AlternateContent after
   * workbookPr) — round-trip only.
   */
  absPath?: string;
  /** Namespace prefix for absPath; explicit true preserves the legacy x15ac spelling. */
  absPathLegacyPrefix?: boolean;
  /** Legacy SpreadsheetML 2005 workbook child order; round-trip only. */
  legacyChildOrder?: boolean;
  /** Coauthoring revision state (xr:revisionPtr) — round-trip only. */
  revisionPtr?: RevisionPtrOptions;
  /** Trailing extension list (workbook > extLst > ext) — round-trip only. */
  extensions?: WorkbookExtensionOptions[];
  /** Microsoft classification-label metadata part. */
  classificationLabels?: ClassificationLabelListOptions;
  /** Source classification-label part path; default docMetadata/LabelInfo.xml. Round-trip only. */
  classificationLabelsPath?: string;
  /** Threaded-comment authors part backing threaded comments. */
  persons?: PersonsOptions;
  /** Source persons part path; default xl/persons/person.xml. Round-trip only. */
  personsPath?: string;
  /** Rich-data parts backing pictures in cells and similar values. */
  richData?: RichDataPackageOptions;
  /** Chart-only sheets (no cells, just a chart) */
  chartsheets?: ChartsheetOptions[];
  /** Legacy Excel 5.0 dialog sheets (xl/dialogSheets/sheetN.xml) */
  dialogsheets?: DialogsheetOptions[];
  /** Pre-defined differential formats for conditional formatting */
  dxfs?: DxfOptions[];
  /**
   * Fonts section of xl/styles.xml, in source order — round-trip only. With
   * fills/borders/cellXfs present, the parsed style table is adopted wholesale
   * (cells keep raw style indices, matching the source's numbering).
   */
  fonts?: FontOptions[];
  /** Fonts container metadata (CT_Fonts `@x14ac:knownFonts`) — round-trip only. */
  fontsContainer?: FontsContainerOptions;
  /** Fills section of xl/styles.xml, in source order — round-trip only. */
  fills?: CellFillOptions[];
  /** Borders section of xl/styles.xml, in source order — round-trip only. */
  borders?: BorderSideOptions[];
  /** cellXfs section of xl/styles.xml, in source order — round-trip only. */
  cellXfs?: IndexedXfEntry[];
  /** numFmts section of xl/styles.xml, in source order — round-trip only. */
  numFmts?: NumFmtEntry[];
  /** Custom color palette (CT_Colors) */
  colors?: ColorsOptions;
  /** Table styles section of xl/styles.xml (CT_TableStyles), including container attributes */
  tableStyles?: TableStylesInfo;
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
  /**
   * Workbook-level custom view snapshots. Excel requires one entry per
   * sheet-level customSheetViews guid (matched by guid).
   */
  customWorkbookViews?: CustomWorkbookViewOptions[];
  /** File recovery properties */
  fileRecovery?: FileRecoveryPropertiesOptions;
  /** Custom VBA function group names */
  functionGroups?: string[];
  /** Web publishing properties */
  webPublishing?: WebPublishingOptions;
  /** File sharing / read-only recommendation */
  fileSharing?: FileSharingOptions;
  /** Volatile function types (xl/volTypes.xml part, CT_VolTypes root) */
  volTypes?: VolTypeOptions[];
  /** Declared volatile-entry count (CT_VolTypes `@count`); round-trip only. */
  volTypesCount?: number;
  /** Source path of the volatile part; older packages use volatileDependencies.xml. */
  volTypesPath?: string;
  /** Web publish objects (CT_WebPublishItems) */
  webPublishObjects?: WebPublishObjectOptions[];
  /** Defined names — named ranges, constants, formulas (CT_DefinedNames) */
  definedNames?: DefinedNameOptions[];
  /** Workbook view (CT_BookView) */
  bookView?: WorkbookViewOptions | false;
  /** Excel version stamp (CT_FileVersion); false preserves source absence */
  fileVersion?: FileVersionOptions | false;
  /** Calculation properties (CT_CalcPr) */
  calculation?: CalculationPropertiesOptions;
  /** OLE embedded range (CT_OleSize) — workbook-level, after calcPr */
  oleSize?: string;
  /** Workbook properties (CT_WorkbookPr) — includes defaultThemeVersion, which Excel
   * requires when the package carries a theme (dropping it makes the file unopenable) */
  properties?: WorkbookPropertiesOptions;
  /** Calculation chain cells (xl/calcChain.xml) — set on parse; the compiler
   * round-trips them verbatim and rebuilds from formula cells only when absent */
  calcChain?: CalcCell[] | false;
  /** Pivot caches, either generated from source data or fully defined. */
  pivotCaches?: PivotCacheOptions[];
  /**
   * Shared string table (xl/sharedStrings.xml), in table order. Round-trip only:
   * parse fills it so rich-text si entries keep their structure and index; fresh
   * documents omit it and strings register from cell values on stringify.
   */
  sharedStrings?: (string | RichTextOptions)[];
  /** Source <sst/@count>: total string-cell references; undefined on fresh authoring. */
  sharedStringsCount?: number;

  /** Declared unique string count (xl/sharedStrings.xml `@uniqueCount`). Round-trip only when absent in source. */
  sharedStringsUniqueCount?: number;
  /** True when the source declared xl/sharedStrings.xml, even with no si entries. Round-trip only. */
  sharedStringsDeclared?: boolean;
  /** Source core-properties part path; preserves non-default OPC layouts. Round-trip only. */
  corePropertiesPath?: string;
  /**
   * Source root-rels core-properties relationship type. Both the
   * `package/2006` and `officeDocument/2006` URI forms appear in real
   * packages; round-trip only.
   */
  corePropertiesRelationshipType?: string;
  /** Source extended-properties part path; preserves non-default OPC layouts. Round-trip only. */
  appPropertiesPath?: string;
  /** Workbook conformance level (CT_Workbook `@conformance`). */
  conformance?: WorkbookConformance;
  /** Source custom-properties part path; preserves non-default OPC layouts. Round-trip only. */
  customPropertiesPath?: string;
  /** True when the source declared a custom-properties part, even if it has no properties. */
  customPropertiesDeclared?: boolean;
  /** Extended properties (docProps/app.xml) */
  appProperties?: AppPropertiesOptions;
  /** Custom properties (docProps/custom.xml); omitted from the package when empty */
  customProperties?: CustomPropertyOptions[];
  /** Shared-workbook revision log (xl/revisionHeaders.xml + xl/revisions/revisionN.xml + xl/users.xml). */
  revisionLog?: SharedWorkbookOptions;
  /**
   * Content types from the source [Content_Types].xml (round-trip only).
   * Present, generate() keeps the source Default/Override entries as the
   * base declaration table and derives only what they leave uncovered.
   */
  contentTypes?: ContentTypesInput;
  /**
   * OPC relationship/content-type namespace flavor. `microsoft2005` preserves
   * legacy Microsoft package manifests; round-trip only.
   */
  relationshipNamespace?: "standard" | "microsoft2005";
  /**
   * Parts carried verbatim from the source that generate() does not rebuild
   * (drawings, VML, external links, unknown parts) — bytes and content type
   * intact; compiler-rebuilt parts under the same path win. Round-trip channel:
   * captured from a parsed source document — do not hand-author.
   */
  rawParts?: { path: string; data: DataType; contentType?: string }[];
  /**
   * Relationships from rebuilt parts' source .rels that point at rawParts
   * (e.g. workbook.xml → externalLink). Re-emitted verbatim with fresh rId.
   * Round-trip channel: captured from a parsed source document — do not
   * hand-author.
   */
  passthroughRelationships?: {
    source: string;
    relationshipType: string;
    target: string;
    rId: string;
    targetMode?: "External";
  }[];
}

/** Generate a pivot cache from worksheet records. */
export interface SourcePivotCacheOptions {
  mode: "source";
  sourceData: PivotSourceData;
  definition?: PivotCacheDefinitionOptions;
}

/** A fully parsed or hand-authored pivot cache pair. */
export interface DefinitionPivotCacheOptions {
  mode: "definition";
  cacheId: number;
  definitionPath: string;
  recordsPath?: string;
  definition: PivotCacheDefinitionOptions;
  records?: PivotCacheRecordsOptions;
}

export type PivotCacheOptions = SourcePivotCacheOptions | DefinitionPivotCacheOptions;

/** Shared-workbook revision tracking data. */
export interface SharedWorkbookOptions {
  /** xl/revisionHeaders.xml (CT_RevisionHeaders). */
  headers: RevisionHeadersOptions;
  /** Revision logs (CT_Revisions), one per header entry. logs[i] maps to headers.headers[i].rId. */
  logs: RevisionLogOptions[];
  /** xl/users.xml (CT_Users), optional. */
  users?: UsersOptions;
  /** Source users relationship URI; preserves legacy `usernames` spelling. */
  usersRelationshipType?: string;
  /**
   * Source path of revisionHeaders.xml relative to xl/ (e.g.
   * "revisions/revisionHeaders.xml"). Round-trip only: fresh authoring writes
   * xl/revisionHeaders.xml.
   */
  headersPath?: string;
  /**
   * Source path of users.xml relative to xl/. Round-trip only: fresh
   * authoring writes xl/users.xml.
   */
  usersPath?: string;
}
