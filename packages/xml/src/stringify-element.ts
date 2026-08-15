/**
 * Serialize an Element including its own opening/closing tag.
 *
 * `stringify` serializes only an element's children (it treats its input as
 * a document root). Raw-XML round-trip of whole elements needs the element's
 * own tag wrapped around its serialized children.
 */
import { escapeXml } from "./escape";
import { stringify as stringifyChildren } from "./stringify";
import type { Element } from "./types";

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
  return `<${el.name}${attrStr}>${stringifyChildren(el)}</${el.name}>`;
}
