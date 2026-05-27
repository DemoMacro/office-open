/**
 * XML Components module — re-exports from core + pptx-specific additions.
 *
 * @module
 */
// Core re-exports
export { BaseXmlComponent } from "./base";
export type { Context, IXmlableObject } from "./base";
export { XmlComponent, IgnoreIfEmptyXmlComponent, EMPTY_OBJECT } from "@office-open/core";
export { XmlAttributeComponent, NextAttributeComponent } from "@office-open/core";
export type { AttributeMap, AttributeData, AttributePayload } from "@office-open/core";
export {
  EmptyElement,
  BuilderElement,
  chartAttr,
  wrapEl,
  stringContainerObj,
} from "@office-open/core";
export {
  ImportedXmlComponent,
  ImportedRootElementAttributes,
  convertToXmlComponent,
} from "@office-open/core";
export { InitializableXmlComponent } from "@office-open/core";
