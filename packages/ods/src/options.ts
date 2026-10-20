import type { PackageOptions } from "@office-open/odf";
import type { WorkbookOptions } from "@office-open/xlsx";

import type { DataMappingOptions } from "./data-mappings";

/** Canonical workbook model plus the unified ODF package overlay. */
export interface OdsWorkbookOptions extends WorkbookOptions, PackageOptions {
  /** LibreOffice external data providers, in source order. */
  dataMappings?: DataMappingOptions[];
}
