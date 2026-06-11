/**
 * Parser for custom XML block elements (w:customXml).
 *
 * @module
 */
import { attr, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import type { SectionChild } from "@shared/section";

import type { DocxReadContext } from "../../context";
import type { CustomXmlBlockOptions, CustomXmlPrOptions } from "./custom-xml";

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
  const opts: Record<string, unknown> = {};

  // Required attribute
  const element = attr(el, "w:element");
  if (element) opts.element = element;

  // Optional URI
  const uri = attr(el, "w:uri");
  if (uri) opts.uri = uri;

  // Parse w:customXmlPr
  const xmlPr = findChild(el, "w:customXmlPr");
  if (xmlPr) {
    opts.customXmlPr = parseCustomXmlPr(xmlPr);
  }

  // Parse block-level children
  const children: SectionChild[] = [];
  for (const child of el.elements ?? []) {
    if (child.name === "w:customXmlPr") continue;
    const parsed = parseChild(child, ctx);
    children.push(parsed);
  }
  if (children.length > 0) opts.children = children;

  return opts as unknown as CustomXmlBlockOptions;
}

function parseCustomXmlPr(el: Element): CustomXmlPrOptions {
  const opts: Record<string, unknown> = {};

  const placeholder = findChild(el, "w:placeholder");
  if (placeholder) {
    const val = attr(placeholder, "w:val");
    if (val) opts.placeholder = val;
  }

  const attributes: { name: string; val: string; uri?: string }[] = [];
  for (const child of el.elements ?? []) {
    if (child.name !== "w:attr") continue;
    const name = attr(child, "w:name");
    const val = attr(child, "w:val");
    if (name && val) {
      const attrOpts: { name: string; val: string; uri?: string } = { name, val };
      const uriVal = attr(child, "w:uri");
      if (uriVal) attrOpts.uri = uriVal;
      attributes.push(attrOpts);
    }
  }
  if (attributes.length > 0) opts.attributes = attributes;

  return opts as unknown as CustomXmlPrOptions;
}
