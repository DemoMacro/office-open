/**
 * Run properties parser for DOCX documents.
 *
 * Parses w:rPr Element trees into RunPropertiesOptions objects.
 *
 * @module
 */
import {
  attr,
  attrBool,
  attrMeasure,
  attrNum,
  colorAttr,
  findChild,
  textOf,
} from "@office-open/xml";
import type { Element } from "@office-open/xml";
import type { RunPropertiesOptions, RunOptions } from "@parts/paragraph/run";

import type { DocxReadContext } from "../../../context";
import { stringifyElement } from "../../../util/stringify-element";
import type { LanguageOptions } from "./language";

/**
 * Parse a w:rPr element into RunPropertiesOptions.
 */
export function parseRunProperties(el: Element): RunPropertiesOptions {
  const opts: Record<string, unknown> = {};

  const rStyle = findChild(el, "w:rStyle");
  if (rStyle) opts.style = attr(rStyle, "w:val");

  const font = findChild(el, "w:rFonts");
  if (font) {
    const ascii = attr(font, "w:ascii");
    const eastAsia = attr(font, "w:eastAsia");
    const hAnsi = attr(font, "w:hAnsi");
    const cs = attr(font, "w:cs");
    const asciiTheme = attr(font, "w:asciiTheme");
    const eastAsiaTheme = attr(font, "w:eastAsiaTheme");
    const hAnsiTheme = attr(font, "w:hAnsiTheme");
    const cstheme = attr(font, "w:cstheme");
    const hint = attr(font, "w:hint");

    if (
      ascii &&
      !eastAsia &&
      !hAnsi &&
      !cs &&
      !asciiTheme &&
      !eastAsiaTheme &&
      !hAnsiTheme &&
      !cstheme
    ) {
      opts.font = hint ? { name: ascii, hint } : ascii;
    } else {
      const fontObj: Record<string, string | undefined> = {};
      if (ascii) fontObj.ascii = ascii;
      if (eastAsia) fontObj.eastAsia = eastAsia;
      if (hAnsi) fontObj.hAnsi = hAnsi;
      if (cs) fontObj.cs = cs;
      if (asciiTheme) fontObj.asciiTheme = asciiTheme;
      if (eastAsiaTheme) fontObj.eastAsiaTheme = eastAsiaTheme;
      if (hAnsiTheme) fontObj.hAnsiTheme = hAnsiTheme;
      if (cstheme) fontObj.cstheme = cstheme;
      if (hint) fontObj.hint = hint;
      opts.font = fontObj;
    }
  }

  const bold = findChild(el, "w:b");
  if (bold) opts.bold = attrBool(bold, "w:val") ?? true;

  const boldCs = findChild(el, "w:bCs");
  if (boldCs) opts.boldComplexScript = attrBool(boldCs, "w:val") ?? true;

  const italic = findChild(el, "w:i");
  if (italic) opts.italic = attrBool(italic, "w:val") ?? true;

  const italicCs = findChild(el, "w:iCs");
  if (italicCs) opts.italicComplexScript = attrBool(italicCs, "w:val") ?? true;

  const underline = findChild(el, "w:u");
  if (underline) {
    const ul: Record<string, string | undefined> = {};
    const uType = attr(underline, "w:val");
    if (uType) ul.type = uType;
    const uColor = colorAttr(underline, "w:color");
    if (uColor) ul.color = uColor;
    opts.underline = ul;
  }

  // On/off properties
  for (const [name, optKey] of [
    ["w:strike", "strike"],
    ["w:dstrike", "doubleStrike"],
    ["w:outline", "outline"],
    ["w:shadow", "shadow"],
    ["w:emboss", "emboss"],
    ["w:imprint", "imprint"],
    ["w:vanish", "vanish"],
    ["w:webHidden", "webHidden"],
    ["w:noProof", "noProof"],
    ["w:snapToGrid", "snapToGrid"],
    ["w:smallCaps", "smallCaps"],
    ["w:caps", "allCaps"],
    ["w:rtl", "rightToLeft"],
    ["w:cs", "complexScript"],
    ["w:specVanish", "specVanish"],
    ["w:oMath", "math"],
  ] as const) {
    const child = findChild(el, name);
    if (child) opts[optKey] = attrBool(child, "w:val") ?? true;
  }

  const color = findChild(el, "w:color");
  if (color) {
    const c = colorAttr(color, "w:val");
    const themeColor = attr(color, "w:themeColor");
    const themeTint = attr(color, "w:themeTint");
    const themeShade = attr(color, "w:themeShade");
    if (themeColor || themeTint || themeShade) {
      const colorObj: Record<string, string | undefined> = {};
      if (c) colorObj.val = c;
      if (themeColor) colorObj.themeColor = themeColor;
      if (themeTint) colorObj.themeTint = themeTint;
      if (themeShade) colorObj.themeShade = themeShade;
      opts.color = colorObj;
    } else if (c) {
      opts.color = c;
    }
  }

  const sz = findChild(el, "w:sz");
  if (sz) {
    const halfPts = attrNum(sz, "w:val");
    if (halfPts !== undefined) opts.size = halfPts / 2;
  }

  const szCs = findChild(el, "w:szCs");
  if (szCs) {
    const halfPts = attrNum(szCs, "w:val");
    if (halfPts !== undefined) opts.sizeComplexScript = halfPts / 2;
  }

  const highlight = findChild(el, "w:highlight");
  if (highlight) {
    const val = attr(highlight, "w:val");
    if (val) opts.highlight = val;
  }

  const highlightCs = findChild(el, "w:highlightCs");
  if (highlightCs) {
    const val = attr(highlightCs, "w:val");
    if (val) opts.highlightComplexScript = val;
  }

  const vertAlign = findChild(el, "w:vertAlign");
  if (vertAlign) {
    const val = attr(vertAlign, "w:val");
    if (val === "subscript") opts.subScript = true;
    else if (val === "superscript") opts.superScript = true;
  }

  const effect = findChild(el, "w:effect");
  if (effect) {
    const val = attr(effect, "w:val");
    if (val) opts.effect = val;
  }

  const emphasisMark = findChild(el, "w:em");
  if (emphasisMark) {
    const val = attr(emphasisMark, "w:val");
    if (val) opts.emphasisMark = { type: val };
  }

  const spacing = findChild(el, "w:spacing");
  if (spacing) {
    const val = attrMeasure(spacing, "w:val");
    if (val !== undefined) opts.characterSpacing = val;
  }

  const scale = findChild(el, "w:w");
  if (scale) {
    const val = attrNum(scale, "w:val");
    if (val !== undefined) opts.scale = val;
  }

  const kern = findChild(el, "w:kern");
  if (kern) {
    const val = attrNum(kern, "w:val");
    if (val !== undefined) opts.kern = val;
  }

  const position = findChild(el, "w:position");
  if (position) {
    const val = attr(position, "w:val");
    if (val !== undefined) opts.position = val;
  }

  const fitText = findChild(el, "w:fitText");
  if (fitText) {
    const val = attrNum(fitText, "w:val");
    if (val !== undefined) opts.fitText = val;
  }

  const lang = findChild(el, "w:lang");
  if (lang) {
    const langObj: LanguageOptions = {};
    const val = attr(lang, "w:val");
    if (val) langObj.value = val;
    const eastAsia = attr(lang, "w:eastAsia");
    if (eastAsia) langObj.eastAsia = eastAsia;
    const bidi = attr(lang, "w:bidi");
    if (bidi) langObj.bidirectional = bidi;
    if (Object.keys(langObj).length > 0) opts.language = langObj;
  }

  // Border (w:bdr)
  const bdr = findChild(el, "w:bdr");
  if (bdr) {
    opts.border = parseBorder(bdr);
  }

  // Shading (w:shd)
  const shd = findChild(el, "w:shd");
  if (shd) {
    opts.shading = parseShading(shd);
  }

  // East Asian layout (w:eastAsianLayout)
  const eastAsianLayout = findChild(el, "w:eastAsianLayout");
  if (eastAsianLayout) {
    opts.eastAsianLayout = parseEastAsianLayout(eastAsianLayout);
  }

  // Content part (w:contentPart)
  const contentPart = findChild(el, "w:contentPart");
  if (contentPart) {
    const rId = attr(contentPart, "r:id");
    if (rId) opts.contentPartRId = rId;
  }

  // Revision (w:rPrChange)
  const rPrChange = findChild(el, "w:rPrChange");
  if (rPrChange) {
    const rev: Record<string, unknown> = {};
    const author = attr(rPrChange, "w:author");
    if (author) rev.author = author;
    const date = attr(rPrChange, "w:date");
    if (date) rev.date = date;
    const id = attrNum(rPrChange, "w:id");
    if (id !== undefined) rev.id = id;
    const innerRPr = findChild(rPrChange, "w:rPr");
    if (innerRPr) {
      Object.assign(rev, parseRunProperties(innerRPr));
    }
    if (Object.keys(rev).length > 0) opts.revision = rev;
  }

  // w14:* text effects (glow/shadow/reflection/props3d) occupy the EG_RPrBase
  // extension slot at the end of rPr. Low-frequency complex subtrees — kept
  // verbatim as raw XML for fidelity while the rPr backbone stays editable.
  const w14Parts: string[] = [];
  for (const child of el.elements ?? []) {
    if (child.name?.startsWith("w14:")) w14Parts.push(stringifyElement(child));
  }
  if (w14Parts.length > 0) opts.w14RawXml = w14Parts.join("");

  return opts as RunPropertiesOptions;
}

/**
 * Parse a w:bdr element into BorderOptions.
 */
export function parseBorder(el: Element): Record<string, unknown> {
  const opts: Record<string, unknown> = {};
  const style = attr(el, "w:val");
  if (style) opts.style = style;
  const color = colorAttr(el, "w:color");
  if (color) opts.color = color;
  const size = attrNum(el, "w:sz");
  if (size !== undefined) opts.size = size;
  const space = attrNum(el, "w:space");
  if (space !== undefined) opts.space = space;
  const shadow = attrBool(el, "w:shadow");
  if (shadow !== undefined) opts.shadow = shadow;
  const frame = attrBool(el, "w:frame");
  if (frame !== undefined) opts.frame = frame;
  return opts;
}

/**
 * Parse a w:shd element into ShadingAttributesProperties.
 */
export function parseShading(el: Element): Record<string, unknown> {
  const opts: Record<string, unknown> = {};
  const fill = colorAttr(el, "w:fill");
  if (fill) opts.fill = fill;
  const color = colorAttr(el, "w:color");
  if (color) opts.color = color;
  const type = attr(el, "w:val");
  if (type) opts.type = type;
  return opts;
}

/**
 * Parse a w:eastAsianLayout element into EastAsianLayoutOptions.
 */
export function parseEastAsianLayout(el: Element): Record<string, unknown> {
  const opts: Record<string, unknown> = {};
  const id = attrNum(el, "w:id");
  if (id !== undefined) opts.id = id;
  const combine = attrBool(el, "w:combine");
  if (combine !== undefined) opts.combine = combine;
  const combineBrackets = attr(el, "w:combineBrackets");
  if (combineBrackets) opts.combineBrackets = combineBrackets;
  const vert = attrBool(el, "w:vert");
  if (vert !== undefined) opts.vert = vert;
  const vertCompress = attrBool(el, "w:vertCompress");
  if (vertCompress !== undefined) opts.vertCompress = vertCompress;
  return opts;
}

// ── Special run child constants ──────────────────────────────────────────────

/** Matches w:br[@w:type="page"] → PageBreak */
export const PARSED_PAGE_BREAK = Symbol("PageBreak");
/** Matches w:br (line break) */
export const PARSED_LINE_BREAK = Symbol("LineBreak");
/** Matches w:tab */
export const PARSED_TAB = Symbol("Tab");
/** Matches w:cr */
export const PARSED_CARRIAGE_RETURN = Symbol("CarriageReturn");
/** Matches w:noBreakHyphen */
export const PARSED_NO_BREAK_HYPHEN = Symbol("NoBreakHyphen");
/** Matches w:softHyphen */
export const PARSED_SOFT_HYPHEN = Symbol("SoftHyphen");
/** Matches w:footnoteRef — auto-generated by Footnote class */
export const PARSED_FOOTNOTE_REF = Symbol("FootnoteRef");
/** Matches w:br[@w:type="column"] */
export const PARSED_COLUMN_BREAK = Symbol("ColumnBreak");
/** Matches w:dayShort */
export const PARSED_DAY_SHORT = Symbol("DayShort");
/** Matches w:monthShort */
export const PARSED_MONTH_SHORT = Symbol("MonthShort");
/** Matches w:yearShort */
export const PARSED_YEAR_SHORT = Symbol("YearShort");
/** Matches w:dayLong */
export const PARSED_DAY_LONG = Symbol("DayLong");
/** Matches w:monthLong */
export const PARSED_MONTH_LONG = Symbol("MonthLong");
/** Matches w:yearLong */
export const PARSED_YEAR_LONG = Symbol("YearLong");
/** Matches w:annotationRef */
export const PARSED_ANNOTATION_REF = Symbol("AnnotationRef");
/** Matches w:separator */
export const PARSED_SEPARATOR = Symbol("Separator");
/** Matches w:continuationSeparator */
export const PARSED_CONTINUATION_SEPARATOR = Symbol("ContinuationSeparator");
/** Matches w:pgNum */
export const PARSED_PAGE_NUMBER = Symbol("PageNumber");
/** Matches w:lastRenderedPageBreak */
export const PARSED_LAST_RENDERED_PAGE_BREAK = Symbol("LastRenderedPageBreak");

export type ParsedRunChild =
  | string
  | typeof PARSED_PAGE_BREAK
  | typeof PARSED_LINE_BREAK
  | typeof PARSED_TAB
  | typeof PARSED_CARRIAGE_RETURN
  | typeof PARSED_NO_BREAK_HYPHEN
  | typeof PARSED_SOFT_HYPHEN
  | typeof PARSED_FOOTNOTE_REF
  | typeof PARSED_COLUMN_BREAK
  | typeof PARSED_DAY_SHORT
  | typeof PARSED_MONTH_SHORT
  | typeof PARSED_YEAR_SHORT
  | typeof PARSED_DAY_LONG
  | typeof PARSED_MONTH_LONG
  | typeof PARSED_YEAR_LONG
  | typeof PARSED_ANNOTATION_REF
  | typeof PARSED_SEPARATOR
  | typeof PARSED_CONTINUATION_SEPARATOR
  | typeof PARSED_PAGE_NUMBER
  | typeof PARSED_LAST_RENDERED_PAGE_BREAK
  | { commentReference: number };

/**
 * Parse a w:r element into run data.
 * Returns { properties, children } where children are parsed run content items.
 */
export function parseRun(
  el: Element,
  _ctx: DocxReadContext,
): {
  properties: RunPropertiesOptions | undefined;
  children: ParsedRunChild[];
  rsid?: string;
  runPropertiesRsid?: string;
  deletionRsid?: string;
} {
  const rPr = findChild(el, "w:rPr");
  const properties = rPr ? parseRunProperties(rPr) : undefined;
  const children: ParsedRunChild[] = [];
  const rsid = attr(el, "w:rsidR");
  const runPropertiesRsid = attr(el, "w:rsidRPr");
  const deletionRsid = attr(el, "w:rsidDel");

  for (const child of el.elements ?? []) {
    switch (child.name) {
      case "w:rPr":
        // already handled above
        break;
      case "w:t": {
        const preserveSpace = attrBool(child, "xml:space");
        let text = textOf(child);
        if (preserveSpace && text) {
          // keep leading/trailing whitespace
          // textOf already returns the raw text
        }
        children.push(text);
        break;
      }
      case "w:delText": {
        // Deleted text in track changes (same format as w:t)
        const text = textOf(child);
        if (text) children.push(text);
        break;
      }
      case "w:br": {
        const brType = attr(child, "w:type");
        if (brType === "page") {
          children.push(PARSED_PAGE_BREAK);
        } else if (brType === "column") {
          children.push(PARSED_COLUMN_BREAK);
        } else {
          children.push(PARSED_LINE_BREAK);
        }
        break;
      }
      case "w:tab":
        children.push(PARSED_TAB);
        break;
      case "w:cr":
        children.push(PARSED_CARRIAGE_RETURN);
        break;
      case "w:noBreakHyphen":
        children.push(PARSED_NO_BREAK_HYPHEN);
        break;
      case "w:softHyphen":
        children.push(PARSED_SOFT_HYPHEN);
        break;
      case "w:commentReference": {
        const id = attrNum(child, "w:id");
        if (id !== undefined) children.push({ commentReference: id });
        break;
      }
      // Drawing/pict are handled at the paragraph level (parseSectionChild in body.ts)
      // where the drawing is extracted and replaced as a paragraph child.
      case "w:drawing":
      case "w:pict":
        break;
      // Symbol run — extract char and font attributes
      case "w:sym": {
        const charVal = attr(child, "w:char");
        const fontVal = attr(child, "w:font");
        if (charVal) {
          children.push({
            symbolRun: { char: charVal, symbolfont: fontVal ?? "Wingdings" },
          } as unknown as ParsedRunChild);
        }
        break;
      }
      // Footnote/endnote reference — preserve as { footnoteReference: id } / { endnoteReference: id }
      case "w:footnoteReference": {
        const id = attrNum(child, "w:id");
        if (id !== undefined) children.push({ footnoteReference: id } as unknown as ParsedRunChild);
        break;
      }
      case "w:endnoteReference": {
        const id = attrNum(child, "w:id");
        if (id !== undefined) children.push({ endnoteReference: id } as unknown as ParsedRunChild);
        break;
      }
      // Footnote/endnote ref mark inside footnote/endnote content —
      // auto-generated by Footnote/Endnote class, skip to avoid duplication.
      case "w:footnoteRef":
      case "w:endnoteRef":
        children.push(PARSED_FOOTNOTE_REF);
        break;
      // Date/time field elements
      case "w:dayShort":
        children.push(PARSED_DAY_SHORT);
        break;
      case "w:monthShort":
        children.push(PARSED_MONTH_SHORT);
        break;
      case "w:yearShort":
        children.push(PARSED_YEAR_SHORT);
        break;
      case "w:dayLong":
        children.push(PARSED_DAY_LONG);
        break;
      case "w:monthLong":
        children.push(PARSED_MONTH_LONG);
        break;
      case "w:yearLong":
        children.push(PARSED_YEAR_LONG);
        break;
      // Other empty run elements
      case "w:annotationRef":
        children.push(PARSED_ANNOTATION_REF);
        break;
      case "w:separator":
        children.push(PARSED_SEPARATOR);
        break;
      case "w:continuationSeparator":
        children.push(PARSED_CONTINUATION_SEPARATOR);
        break;
      case "w:pgNum":
        children.push(PARSED_PAGE_NUMBER);
        break;
      case "w:lastRenderedPageBreak":
        children.push(PARSED_LAST_RENDERED_PAGE_BREAK);
        break;
      default:
        break;
    }
  }

  return { properties, children, rsid, runPropertiesRsid, deletionRsid };
}

/**
 * Convert parsed run data into an RunOptions suitable for the Document constructor.
 * Simplifies the parsed children into text + break format.
 * If the run contains only a commentReference, returns { commentReference: id } instead.
 * If the run only contains footnoteRef/endnoteRef (auto-generated), returns empty options.
 *
 * When empty run elements (tab, noBreakHyphen, date fields, etc.) are present,
 * uses children[] format to preserve them for round-trip fidelity.
 */

/** Mapping from parse symbols to RunOptions child objects for empty elements. */
const SYMBOL_TO_CHILD = new Map<symbol, Record<string, true>>([
  [PARSED_TAB, { tab: true }],
  [PARSED_CARRIAGE_RETURN, { carriageReturn: true }],
  [PARSED_NO_BREAK_HYPHEN, { noBreakHyphen: true }],
  [PARSED_SOFT_HYPHEN, { softHyphen: true }],
  [PARSED_DAY_SHORT, { dayShort: true }],
  [PARSED_MONTH_SHORT, { monthShort: true }],
  [PARSED_YEAR_SHORT, { yearShort: true }],
  [PARSED_DAY_LONG, { dayLong: true }],
  [PARSED_MONTH_LONG, { monthLong: true }],
  [PARSED_YEAR_LONG, { yearLong: true }],
  [PARSED_ANNOTATION_REF, { annotationRef: true }],
  [PARSED_SEPARATOR, { separator: true }],
  [PARSED_CONTINUATION_SEPARATOR, { continuationSeparator: true }],
  [PARSED_PAGE_NUMBER, { pgNum: true }],
  [PARSED_LAST_RENDERED_PAGE_BREAK, { lastRenderedPageBreak: true }],
]);

export function parsedRunToOptions(
  parsed: ReturnType<typeof parseRun>,
): RunOptions | { commentReference: number } | null {
  // Filter out footnoteRef/endnoteRef symbols (auto-generated by Footnote/Endnote class)
  const contentChildren = parsed.children.filter((c) => c !== PARSED_FOOTNOTE_REF);
  const isOnlyFootnoteRef =
    contentChildren.length === 0 && parsed.children.some((c) => c === PARSED_FOOTNOTE_REF);

  // If the run only contained footnoteRef/endnoteRef (no text, no other content),
  // skip it entirely — the Footnote/Endnote class auto-adds FootnoteRefRun.
  if (isOnlyFootnoteRef) {
    return null;
  }

  const opts: Record<string, unknown> = { ...parsed.properties };
  if (parsed.rsid) opts.rsid = parsed.rsid;
  if (parsed.runPropertiesRsid) opts.runPropertiesRsid = parsed.runPropertiesRsid;
  if (parsed.deletionRsid) opts.deletionRsid = parsed.deletionRsid;

  // Check if this run is a pure reference run (commentReference, footnoteReference, endnoteReference)
  const isRefChild = (c: unknown): c is Record<string, number> =>
    typeof c === "object" &&
    c !== null &&
    ("commentReference" in c || "footnoteReference" in c || "endnoteReference" in c);

  const refChildren = contentChildren.filter(isRefChild);
  const nonRefChildren = contentChildren.filter((c) => !isRefChild(c));

  // If the run is a pure reference run (no text), return it directly.
  // Drop auto-generated rStyle (e.g., "FootnoteReference") since it's implicit.
  if (refChildren.length > 0 && nonRefChildren.length === 0) {
    return refChildren[0] as RunOptions | { commentReference: number };
  }

  // If the run only contains a symbolRun, return it directly
  const symbolIdx = nonRefChildren.findIndex(
    (c) => typeof c === "object" && c !== null && "symbolRun" in c,
  );
  if (symbolIdx >= 0 && nonRefChildren.length === 1 && !parsed.properties) {
    return nonRefChildren[symbolIdx] as unknown as RunOptions;
  }

  // Collect text and breaks
  const textParts: string[] = [];
  let breakCount = 0;
  let hasPageBreak = false;
  let hasColumnBreak = false;
  const extraChildren: Record<string, true>[] = [];

  for (const child of nonRefChildren) {
    if (typeof child === "string") {
      textParts.push(child);
    } else if (child === PARSED_LINE_BREAK) {
      breakCount++;
    } else if (child === PARSED_PAGE_BREAK) {
      hasPageBreak = true;
    } else if (child === PARSED_COLUMN_BREAK) {
      hasColumnBreak = true;
    } else {
      // Empty run elements (tab, noBreakHyphen, date fields, etc.)
      const mapped = SYMBOL_TO_CHILD.get(child as symbol);
      if (mapped) extraChildren.push(mapped);
    }
  }

  // When empty elements are present, use children[] format for round-trip fidelity
  if (extraChildren.length > 0) {
    const children: (string | Record<string, unknown>)[] = [];
    for (const child of nonRefChildren) {
      if (typeof child === "string") {
        children.push(child);
      } else if (child === PARSED_LINE_BREAK) {
        children.push({ break: 1 });
      } else if (child === PARSED_PAGE_BREAK) {
        children.push({ pageBreak: true });
      } else if (child === PARSED_COLUMN_BREAK) {
        children.push({ columnBreak: true });
      } else {
        const mapped = SYMBOL_TO_CHILD.get(child as symbol);
        if (mapped) children.push(mapped);
      }
    }
    opts.children = children;
  } else {
    if (textParts.length > 0) {
      opts.text = textParts.join("");
    }
    if (breakCount > 0) {
      opts.break = breakCount;
    }
    if (hasPageBreak) {
      opts.pageBreak = true;
    }
    if (hasColumnBreak) {
      opts.columnBreak = true;
    }
  }

  // If the run has no content and no properties (e.g., a pure drawing run),
  // return null so it can be skipped by the caller.
  if (
    Object.keys(opts).length === 0 &&
    textParts.length === 0 &&
    breakCount === 0 &&
    !hasPageBreak &&
    !hasColumnBreak &&
    extraChildren.length === 0
  ) {
    return null;
  }

  return opts as RunOptions;
}
