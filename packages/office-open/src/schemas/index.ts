/**
 * Schema access: full schemas, validators, entry catalog, slicing, and skeletons.
 *
 * @module
 */
export type { DocumentType, JsonSchema } from "./schemas";
export { docxSchema, pptxSchema, SCHEMAS, validateDocumentInput, xlsxSchema } from "./schemas";
export type { SchemaEntry } from "./entries";
export { entryNames, SCHEMA_ENTRIES, SCHEMA_ROOTS } from "./entries";
export {
  assertKnownDefinitions,
  sliceDocumentSchema,
  sliceSchema,
  UnknownDefinitionError,
} from "./slice";
export { getSkeletonSchema } from "./skeleton";
export { renderSliceTypeText } from "./type-text";
