/**
 * Web Settings module for WordprocessingML documents.
 *
 * Provides types for web settings configuration.
 * XML generation is handled by the descriptor pipeline (webSettingsDesc).
 *
 * Reference: http://officeopenxml.com/WPwebSettings.php
 *
 * @module
 */

import type { FramesetOptions } from "@parts/frameset";

/**
 * Border options for div elements.
 */
export interface DivBorderOptions {
  top?: { style: string; color?: string; size?: number };
  left?: { style: string; color?: string; size?: number };
  bottom?: { style: string; color?: string; size?: number };
  right?: { style: string; color?: string; size?: number };
}

/**
 * Options for a div element (CT_Div).
 */
export interface DivOptions {
  /** Unique div identifier (required by CT_Div/@id) */
  id: number;
  marginLeft: number | UniversalMeasure;
  marginRight: number | UniversalMeasure;
  marginTop: number | UniversalMeasure;
  marginBottom: number | UniversalMeasure;
  /** Mark as HTML blockquote element */
  blockQuote?: boolean;
  /** Mark as HTML body element */
  bodyDiv?: boolean;
  border?: DivBorderOptions;
  children?: DivOptions[];
}

/**
 * Screen size types for web settings target.
 */
export const TargetScreenSize = {
  SIZE_544X376: "544x376",
  SIZE_640X480: "640x480",
  SIZE_720X512: "720x512",
  SIZE_800X600: "800x600",
  SIZE_1024X768: "1024x768",
  SIZE_1152X882: "1152x882",
  SIZE_1152X900: "1152x900",
  SIZE_1280X1024: "1280x1024",
  SIZE_1600X1200: "1600x1200",
  SIZE_1800X1440: "1800x1440",
  SIZE_1920X1200: "1920x1200",
} as const;

/**
 * Options for web settings (CT_WebSettings).
 */
export interface WebSettingsOptions {
  /** Frameset definition for web layout */
  frameset?: FramesetOptions;
  /** Div elements for HTML div formatting */
  divs?: DivOptions[];
  /** Character encoding for web output */
  encoding?: string;
  /** Optimize document rendering for web browser */
  optimizeForBrowser?: boolean;
  /** Rely on VML for graphics display */
  relyOnVML?: boolean;
  /** Allow PNG image format in web output */
  allowPNG?: boolean;
  /** Do not rely on CSS for formatting */
  doNotRelyOnCSS?: boolean;
  /** Do not save as single web file */
  doNotSaveAsSingleFile?: boolean;
  /** Do not organize supporting files in folders */
  doNotOrganizeInFolder?: boolean;
  /** Do not use long file names for supporting files */
  doNotUseLongFileNames?: boolean;
  /** Pixels per inch for web output */
  pixelsPerInch?: number;
  /** Target screen size */
  targetScreenSize?: (typeof TargetScreenSize)[keyof typeof TargetScreenSize] | string;
  /** Save smart tags as XML */
  saveSmartTagsAsXml?: boolean;
}

// ── Descriptor ──

import type { UniversalMeasure } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrBool, attrMeasure, attrNum, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import type { FrameOptions, FramesetSplitbarOptions } from "@parts/frameset";

/** Subset of WebSettingsOptions for descriptor stringify. */
export interface WebSettingsInput {
  frameset?: FramesetOptions;
  divs?: DivOptions[];
  encoding?: string;
  optimizeForBrowser?: boolean;
  relyOnVML?: boolean;
  allowPNG?: boolean;
  doNotRelyOnCSS?: boolean;
  doNotSaveAsSingleFile?: boolean;
  doNotOrganizeInFolder?: boolean;
  doNotUseLongFileNames?: boolean;
  pixelsPerInch?: number;
  targetScreenSize?: string;
  saveSmartTagsAsXml?: boolean;
}

const WS_NS =
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';

function wsOnOff(tag: string, val: boolean): string {
  return `<${tag} w:val="${val ? "true" : "false"}"/>`;
}

function wsStringVal(tag: string, val: string): string {
  return `<${tag} w:val="${wsEscapeAttr(val)}"/>`;
}

function wsNumVal(tag: string, val: number | UniversalMeasure): string {
  return `<${tag} w:val="${val}"/>`;
}

function wsEscapeAttr(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// ── Parse helpers (for descriptor parse path) ──

function parseFramesetEl(el: Element): FramesetOptions {
  const opts: Record<string, unknown> = {};

  const sz = findChild(el, "w:sz");
  if (sz) {
    const val = attr(sz, "w:val");
    if (val) opts.size = val;
  }

  const splitbar = findChild(el, "w:framesetSplitbar");
  if (splitbar) opts.splitbar = parseSplitbarEl(splitbar);

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

  const children: (FramesetOptions | FrameOptions)[] = [];
  for (const child of el.elements ?? []) {
    if (child.name === "w:frameset") {
      children.push(parseFramesetEl(child));
    } else if (child.name === "w:frame") {
      children.push(parseFrameEl(child));
    }
  }
  if (children.length > 0) opts.children = children;

  return opts as unknown as FramesetOptions;
}

function parseSplitbarEl(el: Element): FramesetSplitbarOptions {
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

function parseFrameEl(el: Element): FrameOptions {
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

function parseDivsEl(el: Element): DivOptions[] {
  const result: DivOptions[] = [];
  for (const child of el.elements ?? []) {
    if (child.name === "w:div") {
      result.push(parseDivEl(child));
    }
  }
  return result;
}

function parseDivEl(el: Element): DivOptions {
  const id = attrNum(el, "w:id") ?? 0;
  const opts: Record<string, unknown> = { id };

  const marLeft = findChild(el, "w:marLeft");
  if (marLeft) {
    const val = attrMeasure(marLeft, "w:val");
    if (val !== undefined) opts.marginLeft = val;
  }

  const marRight = findChild(el, "w:marRight");
  if (marRight) {
    const val = attrMeasure(marRight, "w:val");
    if (val !== undefined) opts.marginRight = val;
  }

  const marTop = findChild(el, "w:marTop");
  if (marTop) {
    const val = attrMeasure(marTop, "w:val");
    if (val !== undefined) opts.marginTop = val;
  }

  const marBottom = findChild(el, "w:marBottom");
  if (marBottom) {
    const val = attrMeasure(marBottom, "w:val");
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
    opts.border = parseDivBorderEl(divBdr);
  }

  const divsChild = findChild(el, "w:divsChild");
  if (divsChild) {
    const children = parseDivsEl(divsChild);
    if (children.length > 0) opts.children = children;
  }

  return opts as unknown as DivOptions;
}

function parseDivBorderEl(el: Element): DivBorderOptions {
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

function wsDivBorderXml(b: NonNullable<DivOptions["border"]>): string {
  const parts: string[] = ["<w:divBdr>"];
  const sides: [string, NonNullable<DivOptions["border"]>["top"]][] = [
    ["w:top", b.top],
    ["w:left", b.left],
    ["w:bottom", b.bottom],
    ["w:right", b.right],
  ];
  for (const [tag, side] of sides) {
    if (!side) continue;
    const attrParts: string[] = [`w:val="${wsEscapeAttr(side.style)}"`];
    if (side.color) attrParts.push(`w:color="${wsEscapeAttr(side.color)}"`);
    if (side.size !== undefined) attrParts.push(`w:sz="${side.size}"`);
    parts.push(`<${tag} ${attrParts.join(" ")}/>`);
  }
  parts.push("</w:divBdr>");
  return parts.join("");
}

function wsDivXml(d: DivOptions): string {
  const parts: string[] = [`<w:div w:id="${d.id}">`];
  if (d.blockQuote !== undefined)
    parts.push(wsStringVal("w:blockQuote", d.blockQuote ? "on" : "off"));
  if (d.bodyDiv !== undefined) parts.push(wsStringVal("w:bodyDiv", d.bodyDiv ? "on" : "off"));
  parts.push(wsNumVal("w:marLeft", d.marginLeft));
  parts.push(wsNumVal("w:marRight", d.marginRight));
  parts.push(wsNumVal("w:marTop", d.marginTop));
  parts.push(wsNumVal("w:marBottom", d.marginBottom));
  if (d.border) parts.push(wsDivBorderXml(d.border));
  if (d.children?.length) {
    parts.push("<w:divsChild>");
    for (const child of d.children) parts.push(wsDivXml(child));
    parts.push("</w:divsChild>");
  }
  parts.push("</w:div>");
  return parts.join("");
}

export function framesetXml(fs: FramesetOptions): string {
  const parts: string[] = ["<w:frameset>"];
  if (fs.size !== undefined) parts.push(wsStringVal("w:sz", fs.size));
  if (fs.splitbar) {
    parts.push("<w:framesetSplitbar>");
    if (fs.splitbar.width !== undefined)
      parts.push(`<w:w w:w="${fs.splitbar.width}" w:type="dxa"/>`);
    if (fs.splitbar.color !== undefined) parts.push(wsStringVal("w:color", fs.splitbar.color));
    if (fs.splitbar.noBorder) parts.push("<w:noBorder/>");
    if (fs.splitbar.flatBorders) parts.push("<w:flatBorders/>");
    parts.push("</w:framesetSplitbar>");
  }
  if (fs.layout !== undefined) parts.push(`<w:frameLayout w:val="${fs.layout}"/>`);
  if (fs.title !== undefined) parts.push(wsStringVal("w:title", fs.title));
  if (fs.children) {
    for (const child of fs.children) {
      if ("children" in child || "layout" in child || "splitbar" in child) {
        parts.push(framesetXml(child as FramesetOptions));
      } else {
        parts.push(frameXml(child as FrameOptions));
      }
    }
  }
  parts.push("</w:frameset>");
  return parts.join("");
}

export function frameXml(f: FrameOptions): string {
  const parts: string[] = ["<w:frame>"];
  if (f.size !== undefined) parts.push(wsStringVal("w:sz", f.size));
  if (f.name !== undefined) parts.push(wsStringVal("w:name", f.name));
  if (f.title !== undefined) parts.push(wsStringVal("w:title", f.title));
  if (f.sourceRId !== undefined)
    parts.push(`<w:sourceFileName r:id="${wsEscapeAttr(f.sourceRId)}"/>`);
  if (f.marginWidth !== undefined) parts.push(wsNumVal("w:marW", f.marginWidth));
  if (f.marginHeight !== undefined) parts.push(wsNumVal("w:marH", f.marginHeight));
  if (f.scrollbar !== undefined) parts.push(`<w:scrollbar w:val="${f.scrollbar}"/>`);
  if (f.noResizeAllowed) parts.push("<w:noResizeAllowed/>");
  if (f.linkedToFile) parts.push("<w:linkedToFile/>");
  if (f.longDescRId !== undefined)
    parts.push(`<w:longDesc r:id="${wsEscapeAttr(f.longDescRId)}"/>`);
  parts.push("</w:frame>");
  return parts.join("");
}

export const webSettingsDesc: CustomDescriptor<WebSettingsInput> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const p: string[] = [`<w:webSettings ${WS_NS}>`];

    if (opts.frameset !== undefined) p.push(framesetXml(opts.frameset));
    if (opts.divs?.length) {
      p.push("<w:divs>");
      for (const d of opts.divs) p.push(wsDivXml(d));
      p.push("</w:divs>");
    }
    if (opts.encoding !== undefined) p.push(wsStringVal("w:encoding", opts.encoding));
    if (opts.optimizeForBrowser !== undefined)
      p.push(wsOnOff("w:optimizeForBrowser", opts.optimizeForBrowser));
    if (opts.relyOnVML !== undefined) p.push(wsOnOff("w:relyOnVML", opts.relyOnVML));
    if (opts.allowPNG !== undefined) p.push(wsOnOff("w:allowPNG", opts.allowPNG));
    if (opts.doNotRelyOnCSS !== undefined) p.push(wsOnOff("w:doNotRelyOnCSS", opts.doNotRelyOnCSS));
    if (opts.doNotSaveAsSingleFile !== undefined)
      p.push(wsOnOff("w:doNotSaveAsSingleFile", opts.doNotSaveAsSingleFile));
    if (opts.doNotOrganizeInFolder !== undefined)
      p.push(wsOnOff("w:doNotOrganizeInFolder", opts.doNotOrganizeInFolder));
    if (opts.doNotUseLongFileNames !== undefined)
      p.push(wsOnOff("w:doNotUseLongFileNames", opts.doNotUseLongFileNames));
    if (opts.pixelsPerInch !== undefined) p.push(wsNumVal("w:pixelsPerInch", opts.pixelsPerInch));
    if (opts.targetScreenSize !== undefined)
      p.push(wsStringVal("w:targetScreenSz", opts.targetScreenSize));
    if (opts.saveSmartTagsAsXml !== undefined)
      p.push(wsOnOff("w:saveSmartTagsAsXml", opts.saveSmartTagsAsXml));

    p.push("</w:webSettings>");
    return p.join("");
  },

  parse(el, _ctx) {
    const opts: Record<string, unknown> = {};

    // Frameset
    const frameset = findChild(el, "w:frameset");
    if (frameset) {
      opts.frameset = parseFramesetEl(frameset);
    }

    // Divs
    const divs = findChild(el, "w:divs");
    if (divs) {
      const parsed = parseDivsEl(divs);
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
    const targetScreenSize = findChild(el, "w:targetScreenSz");
    if (targetScreenSize) {
      const val = attr(targetScreenSize, "w:val");
      if (val) opts.targetScreenSize = val;
    }

    return opts as unknown as WebSettingsInput;
  },
};
