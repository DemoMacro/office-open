/**
 * Style cache builder for DOCX parsing.
 *
 * Builds lookup maps from parsed styles.xml and numbering.xml elements
 * for efficient style resolution during document parsing.
 *
 * @module
 */
import { attr, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { DocxDocument } from "../parse";

/**
 * Build a cache of style elements keyed by styleId.
 * Iterates w:styles/w:style elements and indexes by w:styleId/@w:val.
 */
export function buildStyleCache(docx: DocxDocument): Map<string, Element> {
  const cache = new Map<string, Element>();
  if (!docx.styles) return cache;

  for (const child of docx.styles.elements ?? []) {
    if (child.name !== "w:style") continue;
    const styleIdEl = findChild(child, "w:styleId");
    const styleId = styleIdEl ? attr(styleIdEl, "w:val") : undefined;
    if (styleId) {
      cache.set(styleId, child);
    }
  }

  return cache;
}

/**
 * Build a cache of abstract numbering elements keyed by abstractNumId.
 * Iterates w:numbering/w:abstractNum elements and indexes by @w:abstractNumId.
 */
export function buildNumberingCache(docx: DocxDocument): Map<string, Element> {
  const cache = new Map<string, Element>();
  if (!docx.numbering) return cache;

  for (const child of docx.numbering.elements ?? []) {
    if (child.name !== "w:abstractNum") continue;
    const abstractNumId = attr(child, "w:abstractNumId");
    if (abstractNumId !== undefined) {
      cache.set(abstractNumId, child);
    }
  }

  return cache;
}
