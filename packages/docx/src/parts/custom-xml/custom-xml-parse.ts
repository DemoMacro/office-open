/**
 * Parser for custom XML block elements (w:customXml).
 *
 * @module
 */
import { attr, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import type { SectionChild } from "@shared/section";

import type { DocxReadContext } from "../../context";
import { parseCustomXmlProperties } from "../bodychildren";
import type { CustomXmlBlockOptions } from "./custom-xml";

/**
 * Parse w:customXml element into CustomXmlBlockOptions.
 *
 * Uses a callback for child parsing to avoid circular dependencies
 * (same pattern as parseTable).
 */
export function parseCustomXmlBlock(
  el: Element,
  ctx: DocxReadContext,
  parseChild: (el: Element, ctx: DocxReadContext) => SectionChild,
): CustomXmlBlockOptions {
  const opts: Partial<CustomXmlBlockOptions> = {};

  // Required attribute
  const element = attr(el, "w:element");
  if (element) opts.element = element;

  // Optional URI
  const uri = attr(el, "w:uri");
  if (uri) opts.uri = uri;

  // Parse w:customXmlPr
  const xmlPr = findChild(el, "w:customXmlPr");
  if (xmlPr) {
    opts.customXmlPr = parseCustomXmlProperties(xmlPr);
  }

  // Parse block-level children
  const children: SectionChild[] = [];
  for (const child of el.elements ?? []) {
    if (child.name === "w:customXmlPr") continue;
    const parsed = parseChild(child, ctx);
    children.push(parsed);
  }
  if (children.length > 0) opts.children = children;

  return opts as CustomXmlBlockOptions;
}
