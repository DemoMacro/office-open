import type { BaseXmlComponent, Context } from "./xml-components/base";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/**
 * Converts an XmlComponent tree into an XML string via the toXml() fast path.
 */
export class Formatter {
  /**
   * Serialize a component directly to XML string via the toXml() fast path.
   * Always includes XML declaration.
   */
  public formatToXml(input: BaseXmlComponent, context: Context): string {
    const body = input.toXml(context);
    return body ? XML_DECL + body : "";
  }
}
