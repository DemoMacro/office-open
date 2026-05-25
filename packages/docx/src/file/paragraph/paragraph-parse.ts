import { HeadingLevel } from "@file/paragraph/formatting/style";
import type { IParagraphOptions } from "@file/paragraph/paragraph";
/**
 * Paragraph parser for DOCX documents.
 *
 * Parses w:p and w:pPr Element trees into IParagraphOptions objects.
 *
 * @module
 */
import { RawPassthrough } from "@office-open/core";
import { attr, attrBool, attrNum, children, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { ParseContext } from "../../parse/context";
import { parseMathChildren } from "./math/math-parse";
import { parseRun, parseRunProperties, parsedRunToOptions } from "./run/run-parse";

const HEADING_MAP: Record<string, string> = {
  Heading1: HeadingLevel.HEADING_1,
  Heading2: HeadingLevel.HEADING_2,
  Heading3: HeadingLevel.HEADING_3,
  Heading4: HeadingLevel.HEADING_4,
  Heading5: HeadingLevel.HEADING_5,
  Heading6: HeadingLevel.HEADING_6,
  Title: HeadingLevel.TITLE,
};

/**
 * Parse w:pPr element into paragraph properties (without children).
 */
export function parseParagraphProperties(el: Element, _ctx: ParseContext): Record<string, unknown> {
  const opts: Record<string, unknown> = {};

  // Style / heading
  const pStyle = findChild(el, "w:pStyle");
  if (pStyle) {
    const styleVal = attr(pStyle, "w:val");
    if (styleVal) {
      if (HEADING_MAP[styleVal]) {
        opts.heading = HEADING_MAP[styleVal];
      } else {
        opts.style = styleVal;
      }
    }
  }

  // Alignment
  const jc = findChild(el, "w:jc");
  if (jc) {
    const val = attr(jc, "w:val");
    if (val) opts.alignment = val;
  }

  // Spacing
  const spacing = findChild(el, "w:spacing");
  if (spacing) {
    const sp: Record<string, unknown> = {};
    const before = attrNum(spacing, "w:before");
    if (before !== undefined) sp.before = before;
    const after = attrNum(spacing, "w:after");
    if (after !== undefined) sp.after = after;
    const line = attrNum(spacing, "w:line");
    if (line !== undefined) sp.line = line;
    const lineRule = attr(spacing, "w:lineRule");
    if (lineRule) sp.lineRule = lineRule;
    const beforeAuto =
      attrBool(spacing, "w:beforeAutospacing") ?? attrBool(spacing, "w:beforeLines");
    if (beforeAuto !== undefined) sp.beforeAutoSpacing = beforeAuto;
    const afterAuto = attrBool(spacing, "w:afterAutospacing") ?? attrBool(spacing, "w:afterLines");
    if (afterAuto !== undefined) sp.afterAutoSpacing = afterAuto;
    if (Object.keys(sp).length > 0) opts.spacing = sp;
  }

  // Indent
  const ind = findChild(el, "w:ind");
  if (ind) {
    const indentObj: Record<string, unknown> = {};
    const left = attrNum(ind, "w:left");
    if (left !== undefined) indentObj.left = left;
    const right = attrNum(ind, "w:right");
    if (right !== undefined) indentObj.right = right;
    const start = attrNum(ind, "w:start");
    if (start !== undefined) indentObj.start = start;
    const end = attrNum(ind, "w:end");
    if (end !== undefined) indentObj.end = end;
    const hanging = attrNum(ind, "w:hanging");
    if (hanging !== undefined) indentObj.hanging = hanging;
    const firstLine = attrNum(ind, "w:firstLine");
    if (firstLine !== undefined) indentObj.firstLine = firstLine;
    if (Object.keys(indentObj).length > 0) opts.indent = indentObj;
  }

  // Numbering (w:numPr) → bullet
  const numPr = findChild(el, "w:numPr");
  if (numPr) {
    const ilvl = findChild(numPr, "w:ilvl");
    const level = ilvl ? (attrNum(ilvl, "w:val") ?? 0) : 0;
    opts.bullet = { level };
  }

  // Tab stops
  const tabs = findChild(el, "w:tabs");
  if (tabs) {
    const tabStops: Record<string, unknown>[] = [];
    for (const tab of children(tabs, "w:tab")) {
      const tabObj: Record<string, unknown> = {};
      const pos = attrNum(tab, "w:pos");
      if (pos !== undefined) tabObj.position = pos;
      const val = attr(tab, "w:val");
      if (val) tabObj.type = val;
      const leader = attr(tab, "w:leader");
      if (leader) tabObj.leader = leader;
      tabStops.push(tabObj);
    }
    if (tabStops.length > 0) opts.tabStops = tabStops;
  }

  // On/off properties
  for (const [name, optKey] of [
    ["w:keepNext", "keepNext"],
    ["w:keepLines", "keepLines"],
    ["w:pageBreakBefore", "pageBreakBefore"],
    ["w:widowControl", "widowControl"],
    ["w:suppressLineNumbers", "suppressLineNumbers"],
    ["w:contextualSpacing", "contextualSpacing"],
    ["w:bidi", "bidirectional"],
    ["w:wordWrap", "wordWrap"],
    ["w:suppressAutoHyphens", "suppressAutoHyphens"],
    ["w:adjustRightInd", "adjustRightInd"],
    ["w:snapToGrid", "snapToGrid"],
    ["w:mirrorIndents", "mirrorIndents"],
    ["w:kinsoku", "kinsoku"],
    ["w:topLinePunct", "topLinePunct"],
    ["w:autoSpaceDE", "autoSpaceDE"],
    ["w:overflowPunct", "overflowPunctuation"],
    ["w:suppressOverlap", "suppressOverlap"],
  ] as const) {
    const child = findChild(el, name);
    if (child) opts[optKey] = attrBool(child, "w:val") ?? true;
  }

  // Thematic break
  if (findChild(el, "w:pBdr")) {
    const pBdr = findChild(el, "w:pBdr")!;
    const border: Record<string, unknown> = {};
    for (const side of ["top", "bottom", "left", "right"] as const) {
      const sideEl = findChild(pBdr, `w:${side}`);
      if (sideEl) {
        const sideOpts: Record<string, unknown> = {};
        const style = attr(sideEl, "w:val");
        if (style) sideOpts.style = style;
        const color = attr(sideEl, "w:color");
        if (color) sideOpts.color = color;
        const size = attrNum(sideEl, "w:sz");
        if (size !== undefined) sideOpts.size = size;
        const space = attrNum(sideEl, "w:space");
        if (space !== undefined) sideOpts.space = space;
        border[side] = sideOpts;
      }
    }
    if (Object.keys(border).length > 0) opts.border = border;
  }

  // Shading
  const shd = findChild(el, "w:shd");
  if (shd) {
    const shdObj: Record<string, unknown> = {};
    const fill = attr(shd, "w:fill");
    if (fill) shdObj.fill = fill;
    const color = attr(shd, "w:color");
    if (color) shdObj.color = color;
    const val = attr(shd, "w:val");
    if (val) shdObj.type = val;
    if (Object.keys(shdObj).length > 0) opts.shading = shdObj;
  }

  // Text alignment
  const textAlignment = findChild(el, "w:textAlignment");
  if (textAlignment) {
    const val = attr(textAlignment, "w:val");
    if (val) opts.textAlignment = val;
  }

  // Outline level
  const outlineLvl = findChild(el, "w:outlineLvl");
  if (outlineLvl) {
    const val = attrNum(outlineLvl, "w:val");
    if (val !== undefined) opts.outlineLevel = val;
  }

  // Conditional format style
  const cnfStyle = findChild(el, "w:cnfStyle");
  if (cnfStyle) {
    const val = attr(cnfStyle, "w:val");
    if (val) opts.cnfStyle = { val };
  }

  // Run properties (paragraph-level defaults)
  const rPr = findChild(el, "w:rPr");
  if (rPr) {
    opts.run = parseRunProperties(rPr);
  }

  // Frame properties
  const framePr = findChild(el, "w:framePr");
  if (framePr) {
    const frame: Record<string, unknown> = {};
    for (const [attrName, optName] of [
      ["w:dropCap", "dropCap"],
      ["w:lines", "lines"],
      ["w:wrap", "wrap"],
      ["w:vAnchor", "vAnchor"],
      ["w:hAnchor", "hAnchor"],
      ["w:x", "x"],
      ["w:y", "y"],
      ["w:hRule", "hRule"],
      ["w:hSpace", "hSpace"],
      ["w:vSpace", "vSpace"],
    ] as const) {
      const val = attr(framePr, attrName);
      if (val !== undefined) frame[optName] = val;
    }
    const w = attrNum(framePr, "w:w");
    if (w !== undefined) frame.width = w;
    const h = attrNum(framePr, "w:h");
    if (h !== undefined) frame.height = h;
    if (Object.keys(frame).length > 0) opts.frame = frame;
  }

  return opts;
}

/**
 * Parse a w:p element into IParagraphOptions.
 */
export function parseParagraph(el: Element, ctx: ParseContext): IParagraphOptions {
  const opts: Record<string, unknown> = {};

  const pPr = findChild(el, "w:pPr");
  if (pPr) {
    Object.assign(opts, parseParagraphProperties(pPr, ctx));
  }

  // Parse children
  const childList: unknown[] = [];

  for (const child of el.elements ?? []) {
    switch (child.name) {
      case "w:pPr":
        break; // already handled
      case "w:r": {
        const parsed = parseRun(child, ctx);
        // Extract RawPassthrough children and push them directly to paragraph
        const rawChildren = parsed.children.filter(
          (c): c is RawPassthrough => c instanceof RawPassthrough,
        );
        const simplified = {
          ...parsed,
          children: parsed.children.filter((c) => !(c instanceof RawPassthrough)),
        };
        // Convert to IRunOptions (or { commentReference })
        const runOpts = parsedRunToOptions(simplified);
        childList.push(runOpts);
        // Preserve raw passthrough elements for round-trip fidelity
        childList.push(...rawChildren);
        break;
      }
      case "w:hyperlink": {
        // For now, extract text from hyperlink runs
        const linkRuns: unknown[] = [];
        for (const sub of child.elements ?? []) {
          if (sub.name === "w:r") {
            const parsed = parseRun(sub, ctx);
            const runOpts = parsedRunToOptions(parsed);
            linkRuns.push(runOpts);
          }
        }
        // Flatten hyperlink text into the paragraph children
        childList.push(...linkRuns);
        break;
      }
      case "w:bookmarkStart":
      case "w:bookmarkEnd":
        // Bookmarks span across runs and cannot be represented in the current
        // IParagraphOptions JSON output — the text between start/end is already
        // preserved as run children, which is sufficient for most round-trip use cases.
        break;
      case "w:commentRangeStart": {
        const id = attrNum(child, "w:id");
        if (id !== undefined) childList.push({ commentRangeStart: id });
        break;
      }
      case "w:commentRangeEnd": {
        const id = attrNum(child, "w:id");
        if (id !== undefined) childList.push({ commentRangeEnd: id });
        break;
      }
      case "w:commentReference": {
        const id = attrNum(child, "w:id");
        if (id !== undefined) childList.push({ commentReference: id });
        break;
      }
      case "m:oMath": {
        const mathChildren = parseMathChildren(child);
        childList.push({ math: { children: mathChildren } });
        break;
      }
      default:
        // Wrap unknown elements as RawPassthrough
        if (child.name && child.elements && child.elements.length > 0) {
          childList.push(new RawPassthrough(child));
        }
        break;
    }
  }

  // Simple text optimization: if there's only text content, use the text field
  if (childList.length > 0) {
    const allStrings = childList.every(
      (c) =>
        typeof c === "object" &&
        c !== null &&
        "text" in (c as Record<string, unknown>) &&
        Object.keys(c as Record<string, unknown>).length <= 2,
    );
    if (allStrings) {
      const combined = childList.map((c) => (c as Record<string, unknown>).text as string).join("");
      if (combined && Object.keys(opts).length === 0) {
        return combined as unknown as IParagraphOptions;
      }
      if (combined) {
        opts.text = combined;
        return opts as IParagraphOptions;
      }
    }

    opts.children = childList as (import("@file/paragraph/run").IRunOptions | string)[];
  }

  return opts as IParagraphOptions;
}
