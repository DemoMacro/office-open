export { detectOfficeFormat } from "./formats";
export type {
  OfficeFormat,
  OfficeFormatFamily,
  OfficeFormatContainer,
  OfficeFormatInfo,
} from "./formats";
export { detectOfficeFile } from "./detect";

export { generate, generateOfficeDocument, generateToFile, parseInput } from "./generate";
export type {
  GenerateType,
  OfficeGenerateFormat,
  GenerateOptionsMap,
  OutputType,
  GenerateOptions,
} from "./generate";

export { parseOfficeDocument } from "./parse";
export type {
  ParsedOfficeDocument,
  OfficeDocumentParseOptions,
  OfficeWorkbookParseOptions,
  OfficePresentationParseOptions,
} from "./parse";

export { generateDocumentFile, parseDocumentFile } from "./document";
export type {
  DocumentFileGenerateFormat,
  DocumentFileParseFormat,
  DocumentFileParseOptions,
} from "./document";
export { generatePresentationFile, parsePresentationFile } from "./presentation";
export type {
  PresentationFileGenerateFormat,
  PresentationFileParseFormat,
  PresentationFileParseOptions,
} from "./presentation";
export { generateWorkbookFile, parseWorkbookFile } from "./workbook";
export type {
  WorkbookFileGenerateFormat,
  WorkbookFileParseFormat,
  WorkbookFileParseOptions,
} from "./workbook";
