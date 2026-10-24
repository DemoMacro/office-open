/**
 * @office-open/core — Shared OOXML infrastructure.
 *
 * @module
 */
export * from "./smartart";
export * from "./chart";
export * from "./table";
export * from "./picture";
export * from "./connector";
export * from "./group";
export * from "./drawing";
export * from "./vector";
export * from "./patch";
export * from "./theme";
export * from "./util";
export * from "./canonical";

export type { HyperlinkTarget, ReadContext, WriteContext } from "./descriptor";
export { parse as parseDescriptor, stringify as stringifyDescriptor } from "./descriptor";
export { parse, stringify } from "./descriptor";
export type { CustomDescriptor } from "./descriptor";

export {
  convertOutput,
  OoxmlMimeType,
  ooxmlMimeType,
  type OutputByType,
  type OutputType,
  type EncryptedContainerOptions,
  type CompressionOptions,
  type PackerOptions,
  type OoxmlPackageVariant,
} from "./opc";
export {
  createPacker,
  createZipStream,
  ZipStreamWriter,
  zipAndConvert,
  zipSyncAndConvert,
  strFromU8,
  strToU8,
  unzipSync,
  zipSync,
} from "./opc";
export { Media, type BaseMediaEntry } from "./opc";
export { Relationships, type RelationshipType } from "./opc";
export type { CorePropertiesOptions } from "./opc";
export { sanitizeXmlText } from "@office-open/xml";
export { parseArchive, ParsedArchive } from "./opc";
