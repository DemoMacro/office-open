import type { SubDocOptions } from "@file/sub-doc/sub-doc";
/**
 * SubDoc parser for DOCX documents.
 *
 * Parses w:subDoc elements and extracts embedded document data.
 *
 * @module
 */
import { attr } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { ParseContext } from "../../parse/context";

/**
 * Parse a w:subDoc element into SubDocOptions.
 * Reads the referenced document data from the ZIP package.
 */
export function parseSubDoc(el: Element, ctx: ParseContext): SubDocOptions {
  const rId = attr(el, "r:id");
  if (!rId) {
    throw new Error("w:subDoc missing r:id attribute");
  }

  const path = ctx.docx.partRefs.subDocs.get(rId);
  if (!path) {
    throw new Error(`SubDoc relationship ${rId} not found`);
  }

  const data = ctx.docx.doc.getRaw(path);
  if (!data) {
    throw new Error(`SubDoc data not found at ${path}`);
  }

  return { data };
}
