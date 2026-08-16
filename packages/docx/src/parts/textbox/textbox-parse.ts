import { parseVmlShapeStyle, parseVmlStyle } from "@office-open/core";
/**
 * Textbox parser for DOCX documents.
 *
 * Parses w:pict → v:shape → v:textbox → w:txbxContent elements.
 *
 * @module
 */
import { attr, findChild, findFirst } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { DocxReadContext } from "../../context";

/**
 * Parse a w:pict element that contains a textbox.
 * Returns an object suitable for the { textbox: ... } SectionChild variant.
 */
export function parseTextbox(
  el: Element,
  ctx: DocxReadContext,
  parseChildren: (elements: Element[], ctx: DocxReadContext) => unknown[],
): {
  style?: Record<string, string>;
  children?: unknown[];
} {
  const shape = findFirst(el, "v:shape");
  if (!shape) return {};

  const opts: Record<string, unknown> = {};

  // Parse VML style
  const styleAttr = attr(shape, "style");
  if (styleAttr) {
    opts.style = parseVmlShapeStyle(parseVmlStyle(styleAttr));
  }

  // Parse textbox content
  const textbox = findFirst(shape, "v:textbox");
  if (textbox) {
    const txbxContent = findChild(textbox, "w:txbxContent");
    if (txbxContent) {
      const childList = parseChildren(txbxContent.elements ?? [], ctx);
      if (childList.length > 0) opts.children = childList;
    }
  }

  return opts as { style?: Record<string, string>; children?: unknown[] };
}
