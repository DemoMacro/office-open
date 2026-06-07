/**
 * File module — public exports for XLSX document components.
 *
 * @module
 */
export { type WorkbookOptions } from "./file";
export { SharedStrings } from "./shared-strings";
export { Styles } from "./styles";
export type {
  StyleOptions,
  FontOptions,
  FillOptions,
  BorderSideOptions,
  BorderOptions,
  AlignmentOptions,
} from "./styles";
export type {
  WorksheetOptions,
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
  PhoneticPrOptions,
  SheetBackgroundImageOptions,
} from "./worksheet";
export type { IconSetType, CfvoType } from "./worksheet";
export type { CorePropertiesOptions } from "./core-properties";
export type {
  PivotTableOptions,
  PivotDataField,
  ConsolidateFunction,
  PivotFilterOptions,
} from "./pivot";
export { PivotFilterTypeValue } from "./pivot";
export type { TableOptions, TableColumnOptions, TableStyleInfoOptions } from "./table";
export { TotalsRowFunction, TableType } from "./table";

export type {
  WorkbookProtectionOptions,
  CustomWorkbookViewOptions,
  FileRecoveryPrOptions,
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
export type { ConnectionOptions, DbPrOptions, WebPrOptions, ParameterOptions } from "./connection";
export type {
  MapInfoOptions,
  SchemaOptions,
  MapOptions,
  DataBindingOptions,
  XmlPrOptions,
  XmlCellPrOptions,
  SingleXmlCellOptions,
  XmlColumnPrOptions,
} from "./xml-mapping";
