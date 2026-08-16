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
  CellFillOptions,
  BorderSideOptions,
  BorderOptions,
  AlignmentOptions,
  CustomTableStyleOptions,
  CustomCellStyleOptions,
  CellStyleXfOptions,
  StyleExtensionOptions,
} from "./styles";
export { worksheetDesc, stringifyWorksheet, buildWorksheetXml } from "./worksheet";
export type {
  WorksheetOptions,
  PictureOptions,
  WorksheetContext,
  RowOptions,
  CellOptions,
  ColumnOptions,
  MergeCellOptions,
  FreezePaneOptions,
  SheetProtectionOptions,
  ProtectedRangeOptions,
  WorksheetChartOptions,
  AutoFilterOptions,
  FilterColumnOptions,
  Top10FilterOptions,
  CustomFiltersOptions,
  CustomFilterEntry,
  CustomFilterOperator,
  FilterItemsOptions,
  DateGroupFilterOptions,
  DynamicFilterOptions,
  ColorFilterOptions,
  IconFilterOptions,
  SortCondition,
  SortStateOptions,
  HyperlinkOptions,
  CommentOptions,
  NoteAnchorOptions,
  SheetViewOptions,
  PageMarginsOptions,
  PageSetupOptions,
  PrintOptions,
  HeaderFooterOptions,
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
  pickAnchorOptions,
  type DrawingOptions,
  type DrawingPictureOptions,
  type DrawingChartOptions,
  type ShapeOptions,
  type ConnectorOptions,
  type GroupOptions,
  type DrawingContentPartOptions,
  type GroupShapeChildOptions,
  type GroupConnectorChildOptions,
  type DrawingAnchorOptions,
  type AnchorType,
  type EditAsType,
} from "./drawing";
export { externalLinkDesc } from "./external-link";
export { connectionsDesc } from "./connection";
export type {
  ConnectionsOptions,
  ConnectionOptions,
  DatabasePropertiesOptions,
  WebPropertiesOptions,
  TextPropertiesOptions,
  ConnectionTextFieldOptions,
  ParameterOptions,
  WebTableSelection,
} from "./connection";
export { queryTableDesc } from "./query-table";
export type {
  QueryTableOptions,
  QueryTableRefreshOptions,
  QueryTableFieldOptions,
  QueryTableDeletedFieldOptions,
} from "./query-table";
export { metadataDesc } from "./metadata";
export type {
  MetadataOptions,
  MetadataTypeOptions,
  MetadataStringOptions,
  MetadataStringIndexOptions,
  MdxOptions,
  MdxTupleOptions,
  MdxSetOptions,
  MdxMemberPropOptions,
  MdxKpiOptions,
  MdxFunctionType,
  MdxKpiProperty,
  FutureMetadataOptions,
  FutureMetadataBlockOptions,
  MetadataBlockOptions,
  MetadataRecordOptions,
} from "./metadata";
export { pivotTableDesc, type PivotTableDescriptorOptions } from "./pivot-table";
export {
  pivotCacheDefDesc,
  pivotCacheRecordsDesc,
  type PivotCacheDefDescriptorOptions,
} from "./pivot-cache";
export type {
  PivotTableOptions,
  PivotDataField,
  PivotPageFieldOptions,
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
  DefinedNameOptions,
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
export { dialogsheetDesc } from "./dialogsheet";
export type {
  DialogsheetOptions,
  DialogsheetPageMargins,
  DialogsheetPageSetup,
  DialogsheetProtectionOptions,
} from "./dialogsheet";
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
export { mapInfoDesc, singleXmlCellsDesc } from "./xml-mapping";
export type {
  MapInfoOptions,
  SchemaOptions,
  MapOptions,
  DataBindingOptions,
  SingleXmlCellsOptions,
  SingleXmlCellOptions,
  XmlPropertiesOptions,
  XmlCellPropertiesOptions,
  XmlColumnPropertiesOptions,
} from "./xml-mapping";
