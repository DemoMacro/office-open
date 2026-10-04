import type { WorkbookOptions } from "@office-open/xlsx";

import { parseWorkbook as parseWorkbookData } from "./mappers/workbook";

export interface XlsParseOptions {
  /** Password used to verify and decrypt legacy Office RC4 containers. */
  password?: string;
}

export function parseWorkbook(data: Uint8Array, options?: XlsParseOptions): WorkbookOptions {
  return parseWorkbookData(data, options);
}
