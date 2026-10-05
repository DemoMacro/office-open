export {
  detectOffice,
  type OfficeFormat,
  type OfficeFormatContainer,
  type OfficeFormatFamily,
  type OfficeFormatInfo,
} from "./formats";
export type { OutputByType, OutputType } from "@office-open/core";

export {
  generateOffice,
  type OfficeGenerateFormat,
  type OfficeGenerationOptions,
  type OfficeOptionsFor,
} from "./generate";
export { parseOffice, type OfficeParseOptions, type ParsedOffice } from "./parse";

export {
  generateDocument,
  parseDocument,
  patchDocument,
  type DocumentGenerateFormat,
  type DocumentParseFormat,
  type DocumentParseOptions,
  type DocumentPatchFormat,
  type DocumentPatchRequest,
} from "./document";
export {
  generatePresentation,
  parsePresentation,
  patchPresentation,
  type PresentationGenerateFormat,
  type PresentationParseFormat,
  type PresentationParseOptions,
  type PresentationPatchFormat,
  type PresentationPatchRequest,
} from "./presentation";
export {
  generateWorkbook,
  parseWorkbook,
  patchWorkbook,
  type WorkbookGenerateFormat,
  type WorkbookParseFormat,
  type WorkbookParseOptions,
  type WorkbookPatchFormat,
  type WorkbookPatchRequest,
} from "./workbook";
