export {
  generateDocument,
  generateDocumentStream,
  generateDocumentSync,
  parseDocument,
  parseDocumentSync,
} from "./body";
export { patchDocument, type OdtPatchCommand, type OdtPatchOptions } from "./patch";
export { OdtGenerateError, OdtParseError } from "./error";
export type { OdtDocumentOptions } from "./options";
export type {
  OdtEmbeddedObjectOptions,
  OdtEmbeddedObjectFrameOptions,
  OdtFrameAnchorType,
} from "./embedded-objects";
export type { PackageMemberOptions } from "@office-open/odf";
