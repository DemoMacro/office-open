export {
  generateDocument,
  generateDocumentStream,
  generateDocumentSync,
  parseDocument,
  parseDocumentSync,
} from "./body";
export { patchDocument, type OdtPatchCommand, type OdtPatchOptions } from "./patch";
export { OdtGenerateError, OdtParseError } from "./error";
export type { OdtDocumentOptions, OdtVariableDeclarationOptions } from "./options";
export type {
  OdtEmbeddedObjectOptions,
  OdtEmbeddedObjectFrameOptions,
  OdtFrameAnchorType,
} from "./embedded-objects";
export type { OdtCalculationSettingsOptions } from "./options";
export type { OdtIterationOptions, OdtNullDateOptions } from "./calculation-settings";
export type { OdtVariableSetOptions, OdtVariableValueType } from "./paragraph";
export type { PackageMemberOptions } from "@office-open/odf";
