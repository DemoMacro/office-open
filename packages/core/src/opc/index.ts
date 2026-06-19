export { Relationships, TargetModeType } from "./relationships";
export type { RelationshipType } from "./relationships";
export { createDefault, createOverride } from "./content-types";
export type { DefaultAttributes, OverrideAttributes } from "./content-types";

// Core properties (OPC metadata)
export {
  parseCorePropsElement,
  buildCorePropertiesXml,
  buildCorePropertiesXmlString,
  type CorePropertiesOptions,
} from "./core";

// Extended (app) + custom properties (shared OPC parts)
export {
  appPropertiesDesc,
  type AppPropertiesOptions,
  type AppPropertiesInput,
} from "./app-properties";
export {
  customPropertiesDesc,
  type CustomPropertyOptions,
  type CustomPropertiesInput,
} from "./custom-properties";

// Output types
export { convertOutput, OoxmlMimeType, type OutputByType, type OutputType } from "./output";

// ZIP packer
export {
  createPacker,
  createZipStream,
  zipAndConvert,
  zipSyncAndConvert,
  ZIP_DEFLATE_LEVEL,
  ZIP_STORED_LEVEL,
  levelForMediaName,
  type CompileFn,
  type CompressionOptions,
  type Packer,
  type PackerOptions,
  type XmlifyedFile,
} from "./packer";
export type { Zippable, ZipOptions } from "./packer";
export { strFromU8, unzipSync } from "./packer";
export { decodeBase64, encodeBase64 } from "../util/base64";

// Archive parser
export { parseArchive, ParsedArchive } from "./parser";

// OPC consistency validation
export {
  validateOpcConsistency,
  summarizeOpcIssues,
  type OpcIssue,
  type OpcCode,
  type OpcSeverity,
} from "./opc-consistency";
export {
  DOCX_PARTS,
  PPTX_PARTS,
  XLSX_PARTS,
  PART_REGISTRIES,
  type PartDef,
  type PartPresence,
  type PackagePartRegistry,
} from "./part-registry";
export { buildContentTypeOverrides, type ContentTypeOverrideEntry } from "./content-type-overrides";
