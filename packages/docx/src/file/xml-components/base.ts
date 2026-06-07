import type { Context as CoreContext } from "@office-open/core";

/**
 * Docx-specific XML Component bridge.
 *
 * Re-exports BaseXmlComponent from core and defines the docx-specific Context
 * that extends core's generic context with typed file and viewWrapper access.
 *
 * @module
 */
import type { ViewWrapper } from "../document-wrapper";
import type { File } from "../file";

export type { IXmlableObject } from "@office-open/core";
export { BaseXmlComponent } from "@office-open/core";

/**
 * Docx-specific serialization context.
 *
 * Extends core's generic Context with typed file and viewWrapper access.
 * The `file` field is a convenience alias for `fileData`.
 */
export interface Context extends CoreContext<File> {
  /** The root File object being serialized. */
  file: File;
  /** Access to document relationships and other document parts. */
  viewWrapper: ViewWrapper;
}
