/**
 * Element-to-XML serialization helpers shared across the parse layer.
 *
 * @module
 */
import { escapeXml, stringify } from "@office-open/xml";
import type { Element } from "@office-open/xml";

/**
 * Serialize an Element including its own opening/closing tag.
 *
 * `stringify` from @office-open/xml serializes only an element's children
 * (it treats its input as a document root). Raw-XML round-trip of whole
 * elements (TOC paragraphs, range markers, w14 text effects) needs the
 * element's own tag wrapped around its serialized children.
 */
export function stringifyElement(el: Element): string {
  if (!el.name) return "";
  let attrStr = "";
  if (el.attributes) {
    for (const key of Object.keys(el.attributes)) {
      const v = el.attributes[key];
      if (v === null || v === undefined) continue;
      attrStr += ` ${key}="${escapeXml(String(v))}"`;
    }
  }
  const withClosingTag =
    (el.elements?.length ?? 0) > 0 || el.attributes?.["xml:space"] === "preserve";
  if (!withClosingTag) return `<${el.name}${attrStr}/>`;
  return `<${el.name}${attrStr}>${stringify(el)}</${el.name}>`;
}
