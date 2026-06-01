import type { Element } from "@office-open/xml";

import type { Context, IXmlableObject } from "./xml-components";
import { BaseXmlComponent } from "./xml-components";

/**
 * Convert an Element tree to xml-js compact format (IXmlableObject).
 * Always wraps in a named key so the element name is preserved even for empty elements.
 */
export function elementToCompact(el: Element): IXmlableObject {
  const inner: Record<string, unknown> = {};

  if (el.attributes && Object.keys(el.attributes).length > 0) {
    inner._attributes = el.attributes;
  }

  if (el.cdata) {
    inner._cdata = el.cdata;
  } else if (el.text !== null && el.text !== undefined) {
    inner._text = el.text;
  }

  if (el.elements) {
    for (const child of el.elements) {
      const key = child.name ?? "_unknown";
      const compactChild = elementToCompact(child);
      if (inner[key]) {
        if (!Array.isArray(inner[key])) {
          inner[key] = [inner[key]];
        }
        (inner[key] as unknown[]).push(compactChild);
      } else {
        inner[key] = compactChild;
      }
    }
  }

  const name = el.name ?? "unknown";
  // For empty elements, return { name: {} } so serialize produces <name/>
  if (Object.keys(inner).length === 0) {
    return { [name]: {} };
  }
  return { [name]: inner };
}

/**
 * Thin wrapper that passes a raw Element tree through the XML serialization pipeline.
 * Used to include parsed-but-unrecognized XML in generated documents.
 */
export class RawPassthrough extends BaseXmlComponent {
  public constructor(private readonly element: Element) {
    super(element.name ?? "unknown");
  }

  public prepForXml(_context: Context): IXmlableObject {
    return elementToCompact(this.element);
  }
}
