/**
 * Parts module — public exports for XLSX document components.
 *
 * @module
 */
export { type WorkbookOptions } from "./file";
export { SharedStrings, sharedStringsDesc, type SharedStringsDocOptions } from "./shared-strings";
export { Styles, stylesDesc, type StylesDocOptions } from "./styles";
export type {
  StyleOptions,
  FontOptions,
  FillOptions,
  BorderSideOptions,
  BorderOptions,
  AlignmentOptions,
  CustomTableStyleOptions,
  CustomCellStyleOptions,
  StyleExtensionOptions,
} from "./styles";
export { worksheetDesc, stringifyWorksheet, buildWorksheetXml } from "./worksheet";
export type {
  WorksheetOptions,
  WorksheetContext,
  RowOptions,
  CellOptions,
  ColumnOptions,
  MergeCellOptions,
  FreezePaneOptions,
  SheetProtectionOptions,
  ProtectedRangeOptions,
  WorksheetChartOptions,
  DataValidationOptions,
  ConditionalFormatOptions,
  ConditionalFormatRule,
  ColorScaleOptions,
  DataBarOptions,
  IconSetOptions,
  CfvoOptions,
  FormulaOptions,
  ScenarioOptions,
  ScenarioDefinition,
  ScenarioCellOptions,
  IgnoredErrorOptions,
  PhoneticPropertiesOptions,
  SheetBackgroundImageOptions,
} from "./worksheet";
export type { IconSetType, CfvoType } from "./worksheet";
export type { CorePropertiesOptions } from "@office-open/core";
export { calcChainDesc, type CalcChainOptions } from "./calc-chain";
export { chartsheetDesc, type ChartsheetDescriptorOptions } from "./chartsheet";
export { commentsDesc, vmlNotesDesc, type CommentsDocOptions } from "./comments";
export {
  drawingDesc,
  type DrawingOptions,
  type DrawingImageOptions,
  type DrawingChartOptions,
  type AnchorType,
  type EditAsType,
} from "./drawing";
export { externalLinkDesc } from "./external-link";
export { pivotTableDesc, type PivotTableDescriptorOptions } from "./pivot-table";
export {
  pivotCacheDefDesc,
  pivotCacheRecordsDesc,
  type PivotCacheDefDescriptorOptions,
} from "./pivot-cache";
export type {
  PivotTableOptions,
  PivotDataField,
  ConsolidateFunction,
  PivotFilterOptions,
} from "./pivot";
export { PivotFilterTypeValue } from "./pivot";
export { tableDesc } from "./table";
export type { TableOptions, TableColumnOptions, TableStyleInfoOptions } from "./table";
export { TotalsRowFunction, TableType } from "./table";

export { workbookDesc, buildTablePartsXml, buildExternalReferencesXml } from "./workbook";
export type {
  WorkbookProtectionOptions,
  CustomWorkbookViewOptions,
  FileRecoveryPropertiesOptions,
  WebPublishingOptions,
  FileSharingOptions,
} from "./workbook";
export type {
  ExternalLinkOptions,
  ExternalBookOptions,
  ExternalDefinedNameOptions,
  ExternalSheetDataOptions,
  ExternalRowOptions,
  ExternalCellOptions,
  OleLinkOptions,
  OleItemOptions,
} from "./external-link";
export type {
  ChartsheetOptions,
  ChartsheetPageMargins,
  ChartsheetPageSetup,
  ChartsheetProtectionOptions,
  ChartsheetHeaderFooterOptions,
} from "./chartsheet";
export type { DialogsheetOptions } from "./dialogsheet";
export type { QueryTableOptions } from "./query-table";
export type {
  MetadataOptions,
  MetadataTypeOptions,
  MetadataStringOptions,
  FutureMetadataOptions,
} from "./metadata";
export type {
  RevisionHeadersOptions,
  RevisionHeaderEntry,
  RevisionEntry,
  RevisionRowColumnOptions,
  RevisionCellChangeOptions,
  RevisionMoveOptions,
  RevisionFormattingOptions,
  RevisionInsertSheetOptions,
  RevisionCommentOptions,
  RevisionDefinedNameOptions,
  RevisionLogOptions,
} from "./revision-log";
export type {
  ConnectionOptions,
  DatabasePropertiesOptions,
  WebPropertiesOptions,
  ParameterOptions,
} from "./connection";
export type {
  MapInfoOptions,
  SchemaOptions,
  MapOptions,
  DataBindingOptions,
  XmlPropertiesOptions,
  XmlCellPropertiesOptions,
  SingleXmlCellOptions,
  XmlColumnPropertiesOptions,
} from "./xml-mapping";
