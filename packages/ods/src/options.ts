import type { PackageOptions } from "@office-open/odf";
import type { WorkbookOptions } from "@office-open/xlsx";

/** Canonical workbook model plus the unified ODF package overlay. */
export interface OdsWorkbookOptions extends WorkbookOptions, PackageOptions {}
