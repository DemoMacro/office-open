export { APP_PROPS_XML } from "./app-properties";
export { Relationships, TargetModeType } from "./relationships";
export type { RelationshipType } from "./relationships";
export { createDefault, createOverride } from "./content-types";
export type { DefaultAttributes, OverrideAttributes } from "./content-types";

// Core properties (OPC metadata)
export {
  parseCorePropsElement,
  buildCorePropertiesXml,
  buildCorePropertiesXmlString,
  type CoreProperties,
} from "./core";

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
  type CompileFn,
  type CompressionOptions,
  type Packer,
  type PackerOptions,
  type XmlifyedFile,
} from "./packer";
export type { Zippable, ZipOptions } from "./packer";
export { strFromU8, unzipSync } from "./packer";
export { toUint8Array } from "./packer";
export type { DataType } from "./packer";

// Archive parser
export { parseArchive, ParsedArchive } from "./parser";
