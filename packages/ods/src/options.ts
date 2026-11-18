import type { PackageOptions } from "@office-open/odf";
import type { CellOptions, RowOptions, WorkbookOptions, WorksheetOptions } from "@office-open/xlsx";

import type { DataMappingOptions } from "./data-mappings";

/** ODF Data Pilot aggregation with no canonical OOXML equivalent. */
export interface OdsDataPilotFunctionOptions {
  field: string;
  summaryFunction: string;
  tableName?: string;
}

/** ODF what-if scenario attached to a worksheet table. */
export interface OdsScenarioOptions {
  ranges: string;
  active: boolean;
  displayBorder?: boolean;
  borderColor?: string;
  copyBack?: boolean;
  copyStyles?: boolean;
  copyFormulas?: boolean;
  comment?: string;
  protected?: boolean;
}

/** ODF linked source attached to a worksheet table. */
export interface OdsTableSourceOptions {
  href: string;
  mode?: "copy-all" | "copy-results-only";
  tableName?: string;
  actuate?: "onRequest";
  filterName?: string;
  filterOptions?: string;
  refreshDelay?: string;
}

/** ODF cell projection with a repeat count that exceeds XLSX materialization limits. */
export interface OdsCellOptions extends CellOptions {
  repeat?: number;
}

/** ODF row projection with a repeat count that exceeds XLSX materialization limits. */
export interface OdsRowOptions extends Omit<RowOptions, "cells"> {
  cells?: OdsCellOptions[];
  repeat?: number;
}

/** Canonical worksheet plus ODF table-level features without an OOXML equivalent. */
export interface OdsWorksheetOptions extends Omit<WorksheetOptions, "rows"> {
  /** ODF table title; no direct OOXML worksheet equivalent. */
  tableTitle?: string;
  /** ODF table description; no direct OOXML worksheet equivalent. */
  tableDescription?: string;
  scenario?: OdsScenarioOptions;
  tableSource?: OdsTableSourceOptions;
  rows?: OdsRowOptions[];
}

/** ODF calculation settings that extend the canonical workbook calculation model. */
export interface OdsCalculationSettingsOptions extends NonNullable<WorkbookOptions["calculation"]> {
  /** Preserve an ODF iteration element whose optional status is omitted. */
  iterationPresent?: boolean;
  /** Whether search criteria must match the complete cell contents. */
  searchCriteriaMustApplyToWholeCell?: boolean;
  /** ODF compatibility year used when evaluating two-digit dates. */
  nullYear?: number;
}

/** Canonical workbook model plus the unified ODF package overlay. */
export interface OdsWorkbookOptions
  extends Omit<WorkbookOptions, "calculation" | "worksheets">, PackageOptions {
  calculation?: OdsCalculationSettingsOptions;
  worksheets?: OdsWorksheetOptions[];
  /** LibreOffice external data providers, in source order. */
  dataMappings?: DataMappingOptions[];
  dataPilotFunctions?: OdsDataPilotFunctionOptions[];
}
