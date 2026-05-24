/**
 * Textbox parser for DOCX documents.
 *
 * Parses w:pict → v:shape → v:textbox → w:txbxContent elements.
 *
 * @module
 */
import { attr, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { ParseContext } from "../../parse/context";

// Forward declaration
let _parseSectionChildren: ((elements: Element[], ctx: ParseContext) => unknown[]) | undefined;

export function setTextboxSectionChildrenParser(
  fn: (elements: Element[], ctx: ParseContext) => unknown[],
): void {
  _parseSectionChildren = fn;
}

/**
 * Parse VML shape style string into VmlShapeStyle-like object.
 */
function parseVmlStyle(styleStr: string): Record<string, string> {
  const style: Record<string, string> = {};
  for (const part of styleStr.split(";")) {
    const [key, val] = part.split(":").map((s) => s.trim());
    if (key && val) style[key] = val;
  }
  return style;
}

/**
 * Parse a w:pict element that contains a textbox.
 * Returns an object suitable for the { textbox: ... } SectionChild variant.
 */
export function parseTextbox(
  el: Element,
  ctx: ParseContext,
): {
  style?: Record<string, string>;
  children?: unknown[];
} {
  const shape = findDeep(el, "v:shape")[0];
  if (!shape) return {};

  const opts: Record<string, unknown> = {};

  // Parse VML style
  const styleAttr = attr(shape, "style");
  if (styleAttr) {
    opts.style = parseVmlStyle(styleAttr);
  }

  // Parse textbox content
  const textbox = findDeep(shape, "v:textbox")[0];
  if (textbox) {
    const txbxContent = findChild(textbox, "w:txbxContent");
    if (txbxContent && _parseSectionChildren) {
      const childList = _parseSectionChildren(txbxContent.elements ?? [], ctx);
      if (childList.length > 0) opts.children = childList;
    }
  }

  return opts as { style?: Record<string, string>; children?: unknown[] };
}

// Simple deep finder
function findDeep(parent: Element, name: string): Element[] {
  const result: Element[] = [];
  for (const child of parent.elements ?? []) {
    if (child.name === name) result.push(child);
    result.push(...findDeep(child, name));
  }
  return result;
}
