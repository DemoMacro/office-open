/**
 * XML Components module — re-exports from core + docx-specific additions.
 *
 * @module
 */
// Core re-exports
export { BaseXmlComponent } from "./base";
export type { IContext, IXmlableObject } from "./base";
export { XmlComponent, IgnoreIfEmptyXmlComponent, EMPTY_OBJECT } from "@office-open/core";
export { XmlAttributeComponent, NextAttributeComponent } from "@office-open/core";
export type { AttributeMap, AttributeData, AttributePayload } from "@office-open/core";
export {
    OnOffElement,
    HpsMeasureElement,
    EmptyElement,
    StringValueElement,
    NumberValueElement,
    StringEnumValueElement,
    StringContainer,
    BuilderElement,
    chartAttr,
    wrapEl,
} from "@office-open/core";
export {
    ImportedXmlComponent,
    ImportedRootElementAttributes,
    convertToXmlComponent,
} from "@office-open/core";
export { InitializableXmlComponent } from "@office-open/core";

// Docx-specific
export * from "./attributes";
export { createStringElement } from "./simple-elements";
