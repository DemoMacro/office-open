/**
 * @office-open/xlsx — Generate .xlsx files with JS/TS.
 *
 * @module
 */
export * from "./parts";
export * from "./shared";
export * from "./util";
export type {
  CompressionOptions,
  OutputType,
  OutputByType,
  PackerOptions,
} from "@office-open/core";
export { parseXlsx, parseWorkbook } from "./parse";
export type { XlsxDocument, XlsxPartRefs } from "./parse";
export { patchWorkbook, PatchType } from "./patch";
export type {
  Patch,
  CellPatch,
  PatchWorkbookOptions,
  PatchDocumentOutputType,
  InputDataType,
} from "./patch";
export { compileWorkbook } from "./compiler";
export { XlsxWriteContext, XlsxReadContext } from "./context";
export { generateWorkbook, generateWorkbookSync, generateWorkbookStream } from "./generate";
