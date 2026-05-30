/**
 * XML Components module — re-exports from core + docx-specific additions.
 *
 * @module
 */
// Core re-exports
export { BaseXmlComponent } from "./base";
export type { Context, IXmlableObject } from "./base";
import type { BaseXmlComponent as BaseXmlComp, IXmlableObject as XmlObj } from "./base";
/** Union type for values that can be children of an XML element. */
export type BuilderChild = BaseXmlComp | XmlObj | string;
export { XmlComponent, IgnoreIfEmptyXmlComponent, EMPTY_OBJECT } from "@office-open/core";
export type { AttributePayload } from "@office-open/core";
export {
  EmptyElement,
  BuilderElement,
  chartAttr,
  wrapEl,
  onOffObj,
  stringValObj,
  numberValObj,
  stringEnumValObj,
  hpsMeasureObj,
  stringContainerObj,
  attrObj,
} from "@office-open/core";
export {
  ImportedXmlComponent,
  ImportedRootElementAttributes,
  convertToXmlComponent,
} from "@office-open/core";
export { InitializableXmlComponent } from "@office-open/core";

// Docx-specific
export { createStringElement } from "./simple-elements";
