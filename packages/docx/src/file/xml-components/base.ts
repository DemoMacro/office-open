import type { IContext as CoreContext } from "@office-open/core";

/**
 * Docx-specific XML Component bridge.
 *
 * Re-exports BaseXmlComponent from core and defines the docx-specific IContext
 * that extends core's generic context with typed file and viewWrapper access.
 *
 * @module
 */
import type { IViewWrapper } from "../document-wrapper";
import type { File } from "../file";

export type { IXmlableObject } from "@office-open/core";
export { BaseXmlComponent } from "@office-open/core";

/**
 * Docx-specific serialization context.
 *
 * Extends core's generic IContext with typed file and viewWrapper access.
 * The `file` field is a convenience alias for `fileData`.
 */
export interface IContext extends CoreContext<File> {
    /** The root File object being serialized. */
    readonly file: File;
    /** Access to document relationships and other document parts. */
    readonly viewWrapper: IViewWrapper;
}
