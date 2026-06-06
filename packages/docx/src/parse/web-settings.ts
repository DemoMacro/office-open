/**
 * Parser for web settings (word/webSettings.xml).
 *
 * @module
 */
import { attr, attrBool, attrNum, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type {
  FramesetOptions,
  FramesetSplitbarOptions,
  FrameOptions,
} from "../file/frameset/frameset";
import type {
  WebSettingsOptions,
  DivOptions,
  DivBorderOptions,
} from "../file/web-settings/web-settings";

/**
 * Parse word/webSettings.xml root element into WebSettingsOptions.
 */
export function parseWebSettings(el: Element): WebSettingsOptions {
  const opts: Record<string, unknown> = {};

  // Frameset
  const frameset = findChild(el, "w:frameset");
  if (frameset) {
    opts.frameset = parseFrameset(frameset);
  }

  // Divs
  const divs = findChild(el, "w:divs");
  if (divs) {
    const parsed = parseDivs(divs);
    if (parsed.length > 0) opts.divs = parsed;
  }

  // Encoding
  const encoding = findChild(el, "w:encoding");
  if (encoding) {
    const val = attr(encoding, "w:val");
    if (val) opts.encoding = val;
  }

  // Boolean fields
  for (const [name, optKey] of [
    ["w:optimizeForBrowser", "optimizeForBrowser"],
    ["w:relyOnVML", "relyOnVML"],
    ["w:allowPNG", "allowPNG"],
    ["w:doNotRelyOnCSS", "doNotRelyOnCSS"],
    ["w:doNotSaveAsSingleFile", "doNotSaveAsSingleFile"],
    ["w:doNotOrganizeInFolder", "doNotOrganizeInFolder"],
    ["w:doNotUseLongFileNames", "doNotUseLongFileNames"],
    ["w:saveSmartTagsAsXml", "saveSmartTagsAsXml"],
  ] as const) {
    const child = findChild(el, name);
    if (child) opts[optKey] = attrBool(child, "w:val") ?? true;
  }

  // Pixels per inch
  const ppi = findChild(el, "w:pixelsPerInch");
  if (ppi) {
    const val = attrNum(ppi, "w:val");
    if (val !== undefined) opts.pixelsPerInch = val;
  }

  // Target screen size
  const targetSz = findChild(el, "w:targetScreenSz");
  if (targetSz) {
    const val = attr(targetSz, "w:val");
    if (val) opts.targetScreenSz = val;
  }

  return opts as unknown as WebSettingsOptions;
}

function parseFrameset(el: Element): FramesetOptions {
  const opts: Record<string, unknown> = {};

  const sz = findChild(el, "w:sz");
  if (sz) {
    const val = attr(sz, "w:val");
    if (val) opts.size = val;
  }

  const splitbar = findChild(el, "w:framesetSplitbar");
  if (splitbar) opts.splitbar = parseSplitbar(splitbar);

  const layout = findChild(el, "w:frameLayout");
  if (layout) {
    const val = attr(layout, "w:val");
    if (val === "rows" || val === "cols") opts.layout = val;
  }

  const title = findChild(el, "w:title");
  if (title) {
    const val = attr(title, "w:val");
    if (val) opts.title = val;
  }

  // Children: nested framesets and frames
  const children: (FramesetOptions | FrameOptions)[] = [];
  for (const child of el.elements ?? []) {
    if (child.name === "w:frameset") {
      children.push(parseFrameset(child));
    } else if (child.name === "w:frame") {
      children.push(parseFrame(child));
    }
  }
  if (children.length > 0) opts.children = children;

  return opts as unknown as FramesetOptions;
}

function parseSplitbar(el: Element): FramesetSplitbarOptions {
  const opts: Record<string, unknown> = {};

  const w = findChild(el, "w:w");
  if (w) {
    const val = attrNum(w, "w:w");
    if (val !== undefined) opts.width = val;
  }

  const color = findChild(el, "w:color");
  if (color) {
    const val = attr(color, "w:val");
    if (val) opts.color = val;
  }

  if (findChild(el, "w:noBorder")) opts.noBorder = true;
  if (findChild(el, "w:flatBorders")) opts.flatBorders = true;

  return opts as unknown as FramesetSplitbarOptions;
}

function parseFrame(el: Element): FrameOptions {
  const opts: Record<string, unknown> = {};

  const sz = findChild(el, "w:sz");
  if (sz) {
    const val = attr(sz, "w:val");
    if (val) opts.size = val;
  }

  const name = findChild(el, "w:name");
  if (name) {
    const val = attr(name, "w:val");
    if (val) opts.name = val;
  }

  const title = findChild(el, "w:title");
  if (title) {
    const val = attr(title, "w:val");
    if (val) opts.title = val;
  }

  const source = findChild(el, "w:sourceFileName");
  if (source) {
    const val = attr(source, "r:id");
    if (val) opts.sourceRId = val;
  }

  const marW = findChild(el, "w:marW");
  if (marW) {
    const val = attrNum(marW, "w:val");
    if (val !== undefined) opts.marginWidth = val;
  }

  const marH = findChild(el, "w:marH");
  if (marH) {
    const val = attrNum(marH, "w:val");
    if (val !== undefined) opts.marginHeight = val;
  }

  const scrollbar = findChild(el, "w:scrollbar");
  if (scrollbar) {
    const val = attr(scrollbar, "w:val");
    if (val === "on" || val === "off" || val === "auto") opts.scrollbar = val;
  }

  if (findChild(el, "w:noResizeAllowed")) opts.noResizeAllowed = true;
  if (findChild(el, "w:linkedToFile")) opts.linkedToFile = true;

  const longDesc = findChild(el, "w:longDesc");
  if (longDesc) {
    const val = attr(longDesc, "r:id");
    if (val) opts.longDescRId = val;
  }

  return opts as unknown as FrameOptions;
}

function parseDivs(el: Element): DivOptions[] {
  const result: DivOptions[] = [];
  for (const child of el.elements ?? []) {
    if (child.name === "w:div") {
      result.push(parseDiv(child));
    }
  }
  return result;
}

function parseDiv(el: Element): DivOptions {
  const id = attrNum(el, "w:id") ?? 0;
  const opts: Record<string, unknown> = { id };

  const marLeft = findChild(el, "w:marLeft");
  if (marLeft) {
    const val = attrNum(marLeft, "w:val");
    if (val !== undefined) opts.marginLeft = val;
  }

  const marRight = findChild(el, "w:marRight");
  if (marRight) {
    const val = attrNum(marRight, "w:val");
    if (val !== undefined) opts.marginRight = val;
  }

  const marTop = findChild(el, "w:marTop");
  if (marTop) {
    const val = attrNum(marTop, "w:val");
    if (val !== undefined) opts.marginTop = val;
  }

  const marBottom = findChild(el, "w:marBottom");
  if (marBottom) {
    const val = attrNum(marBottom, "w:val");
    if (val !== undefined) opts.marginBottom = val;
  }

  const blockQuote = findChild(el, "w:blockQuote");
  if (blockQuote) {
    const val = attr(blockQuote, "w:val");
    opts.blockQuote = val !== "off";
  }

  const bodyDiv = findChild(el, "w:bodyDiv");
  if (bodyDiv) {
    const val = attr(bodyDiv, "w:val");
    opts.bodyDiv = val !== "off";
  }

  const divBdr = findChild(el, "w:divBdr");
  if (divBdr) {
    opts.border = parseDivBorder(divBdr);
  }

  const divsChild = findChild(el, "w:divsChild");
  if (divsChild) {
    const children = parseDivs(divsChild);
    if (children.length > 0) opts.children = children;
  }

  return opts as unknown as DivOptions;
}

function parseDivBorder(el: Element): DivBorderOptions {
  const opts: Record<string, unknown> = {};

  for (const side of ["top", "left", "bottom", "right"] as const) {
    const sideEl = findChild(el, `w:${side}`);
    if (sideEl) {
      const b: Record<string, unknown> = {};
      const val = attr(sideEl, "w:val");
      if (val) b.style = val;
      const color = attr(sideEl, "w:color");
      if (color) b.color = color;
      const sz = attrNum(sideEl, "w:sz");
      if (sz !== undefined) b.size = sz;
      opts[side] = b;
    }
  }

  return opts as unknown as DivBorderOptions;
}
