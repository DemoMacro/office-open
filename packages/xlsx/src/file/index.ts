/**
 * File module — public exports for XLSX document components.
 *
 * @module
 */
export { File, type WorkbookOptions } from "./file";
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
} from "./worksheet";
export type { IconSetType, CfvoType } from "./worksheet";
export type { CorePropertiesOptions } from "./core-properties";
export type { PivotTableOptions, PivotDataField, ConsolidateFunction } from "./pivot";
export type { TableOptions, TableColumnOptions, TableStyleInfoOptions } from "./table";
export { TotalsRowFunction, TableType } from "./table";
export type { WorkbookProtectionOptions, CustomWorkbookViewOptions } from "./workbook";
export type {
  ExternalLinkOptions,
  ExternalBookOptions,
  ExternalDefinedNameOptions,
  ExternalSheetDataOptions,
} from "./external-link";
export type { ChartsheetOptions } from "./chartsheet";
