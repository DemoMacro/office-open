/**
 * @office-open/xlsx — Generate .xlsx files with JS/TS.
 *
 * @module
 */
export { File as Workbook } from "./file";
export * from "./file";
export * from "./export";
export * from "./util";
export { parseXlsx, parseWorkbook } from "./parse";
export type { XlsxDocument, XlsxPartRefs } from "./parse";
export { patchWorkbook, PatchType } from "./patcher";
export type {
  IPatch,
  CellPatch,
  PatchWorkbookOptions,
  PatchDocumentOutputType,
  InputDataType,
} from "./patcher";
