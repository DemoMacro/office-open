import type { PackageOptions } from "@office-open/odf";
import type { WorkbookOptions } from "@office-open/xlsx";

import type { DataMappingOptions } from "./data-mappings";

/** ODF Data Pilot aggregation with no canonical OOXML equivalent. */
export interface OdsDataPilotFunctionOptions {
  field: string;
  summaryFunction: string;
  tableName?: string;
}

/** ODF calculation settings that extend the canonical workbook calculation model. */
export interface OdsCalculationSettingsOptions extends NonNullable<WorkbookOptions["calculation"]> {
  /** Whether search criteria must match the complete cell contents. */
  searchCriteriaMustApplyToWholeCell?: boolean;
  /** ODF compatibility year used when evaluating two-digit dates. */
  nullYear?: number;
}

/** Canonical workbook model plus the unified ODF package overlay. */
export interface OdsWorkbookOptions extends Omit<WorkbookOptions, "calculation">, PackageOptions {
  calculation?: OdsCalculationSettingsOptions;
  /** LibreOffice external data providers, in source order. */
  dataMappings?: DataMappingOptions[];
  dataPilotFunctions?: OdsDataPilotFunctionOptions[];
}
