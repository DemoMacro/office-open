/**
 * FileChild module for WordprocessingML documents.
 *
 * FileChild is a marker interface for block-level elements that can appear
 * in the document body, such as paragraphs and tables.
 *
 * @module
 */
import type { BaseXmlComponent } from "@file/xml-components";

/**
 * Marker interface for document body children.
 *
 * FileChild identifies a block-level element that can be added directly
 * to the document body. Examples include Paragraph and Table.
 *
 * Consumers should use `instanceof BaseXmlComponent` with the `fileChild`
 * symbol check when needed, or use this type for parameter typing.
 */
export interface FileChild extends BaseXmlComponent {
    /** Marker property identifying this as a FileChild */
    readonly fileChild: symbol;
}
