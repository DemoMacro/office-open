/**
 * Main entry point for the docx library.
 *
 * This module provides a declarative API for creating .docx files that works
 * in both Node.js and browser environments. The library generates documents
 * compliant with the OOXML (Office Open XML) specification.
 *
 * The primary API is `generateDocument()`, which accepts a `DocumentOptions` object
 * and produces a DOCX file directly — no intermediate class required.
 *
 * @module
 *
 * @example
 * ```typescript
 * import { generateDocument } from "@office-open/docx";
 *
 * const document = await generateDocument({
 *   sections: [{
 *     children: [{
 *       paragraph: {
 *         children: [{ text: "Hello World", bold: true }],
 *       },
 *     }],
 *   }],
 * });
 * ```
 */

export * from "./parts";
export * from "./shared";
export * from "./patch";
export type {
  DocumentPatchCommand,
  DocumentPatchField,
  DocumentPatchInsertValue,
  DocumentPatchOptions,
  DocumentPatchValue,
} from "./patch";
export * from "./parse";
export { DocxParseError } from "./errors";
export { generateDocument, generateDocumentSync, generateDocumentStream } from "./generate";
export { compileDocument } from "./compiler";
export { DocxWriteContext, DocxReadContext } from "./context";
export type {
  CompressionOptions,
  OutputType,
  OutputByType,
  PackerOptions,
} from "@office-open/core";
export { framesetXml, frameXml } from "./parts/web-settings";
