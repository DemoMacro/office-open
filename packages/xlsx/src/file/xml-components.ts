/**
 * XML Components module — re-exports from core.
 *
 * @module
 */
export { BaseXmlComponent } from "@office-open/core";
export type { Context, IXmlableObject } from "@office-open/core";
export { XmlComponent, IgnoreIfEmptyXmlComponent, EMPTY_OBJECT } from "@office-open/core";
export type { AttributePayload } from "@office-open/core";
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
