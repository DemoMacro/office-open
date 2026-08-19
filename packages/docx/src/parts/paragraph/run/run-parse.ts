import type { UniversalMeasure } from "@office-open/core";
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
import { objectDesc } from "@parts/object";
import type { ObjectElementOptions } from "@parts/object";
import type { FootnoteEndnoteReferenceOptions } from "@parts/paragraph/paragraph";
import type {
  BreakClear,
  BreakOptions,
  RunPropertiesOptions,
  RunOptions,
} from "@parts/paragraph/run";
import { parsePict } from "@parts/pict";
import type { PictOptions } from "@parts/pict";
import { parseShading } from "@shared/shading";

import type { DocxReadContext } from "../../../context";
import { stringifyElement } from "../../../util/stringify-element";
import type { LanguageOptions } from "./language";

// On/off run properties: XML child tag → options key.
const ON_OFF_RUN_PROPS: readonly (readonly [string, keyof RunPropertiesOptions & string])[] = [
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
];

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
    const complexScript = attr(font, "w:cs");
    const asciiTheme = attr(font, "w:asciiTheme");
    const eastAsiaTheme = attr(font, "w:eastAsiaTheme");
    const hAnsiTheme = attr(font, "w:hAnsiTheme");
    const cstheme = attr(font, "w:cstheme");
    const hint = attr(font, "w:hint");

    if (
      ascii &&
      !eastAsia &&
      !hAnsi &&
      !complexScript &&
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
      if (complexScript) fontObj.complexScript = complexScript;
      if (asciiTheme) fontObj.asciiTheme = asciiTheme;
      if (eastAsiaTheme) fontObj.eastAsiaTheme = eastAsiaTheme;
      if (hAnsiTheme) fontObj.hAnsiTheme = hAnsiTheme;
      if (cstheme) fontObj.complexScriptTheme = cstheme;
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
  for (const [name, optKey] of ON_OFF_RUN_PROPS) {
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

  const vertAlign = findChild(el, "w:vertAlign");
  if (vertAlign) {
    const val = attr(vertAlign, "w:val");
    if (val === "baseline" || val === "subscript" || val === "superscript") {
      opts.verticalAlign = val;
    }
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
    // w:kern is ST_HpsMeasure: numeric tokens are half-points → points (÷2),
    // UniversalMeasure strings pass through verbatim (same split as w:position).
    const val = attrMeasure(kern, "w:val");
    if (val !== undefined) {
      opts.kern = typeof val === "number" ? val / 2 : (val as UniversalMeasure);
    }
  }

  const position = findChild(el, "w:position");
  if (position) {
    // w:position is ST_SignedHpsMeasure; numeric tokens are half-points →
    // points (÷2), UniversalMeasure strings pass through verbatim.
    const val = attrMeasure(position, "w:val");
    if (val !== undefined) {
      opts.position = typeof val === "number" ? val / 2 : (val as UniversalMeasure);
    }
  }

  const fitText = findChild(el, "w:fitText");
  if (fitText) {
    const val = attrNum(fitText, "w:val");
    if (val !== undefined) opts.fitText = val;
  }

  const lang = findChild(el, "w:lang");
  if (lang) {
    // Keep the element even when it carries no attributes — Word writes a
    // bare <w:lang/> to override inherited language settings.
    const langObj: LanguageOptions = {};
    const val = attr(lang, "w:val");
    if (val) langObj.value = val;
    const eastAsia = attr(lang, "w:eastAsia");
    if (eastAsia) langObj.eastAsia = eastAsia;
    const bidi = attr(lang, "w:bidi");
    if (bidi) langObj.bidirectional = bidi;
    opts.language = langObj;
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
  if (vert !== undefined) opts.vertical = vert;
  const vertCompress = attrBool(el, "w:vertCompress");
  if (vertCompress !== undefined) opts.verticalCompress = vertCompress;
  return opts;
}

// ── Special run child constants ──────────────────────────────────────────────

/** Matches w:br[`@w:type`="page"] → PageBreak */
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
/** Matches w:footnoteRef — the reference mark inside footnote content */
export const PARSED_FOOTNOTE_REF = Symbol("FootnoteRef");
/** Matches w:endnoteRef — the reference mark inside endnote content */
export const PARSED_ENDNOTE_REF = Symbol("EndnoteRef");
/** Matches w:br[`@w:type`="column"] */
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
  | typeof PARSED_ENDNOTE_REF
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
  | { commentReference: number }
  | { object: ObjectElementOptions }
  | { break: number | BreakOptions }
  | { footnoteReference: number | FootnoteEndnoteReferenceOptions }
  | { endnoteReference: number | FootnoteEndnoteReferenceOptions };

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
  additionRsid?: string;
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
        const brClear = attr(child, "w:clear");
        if (brType === "page") {
          children.push(PARSED_PAGE_BREAK);
        } else if (brType === "column") {
          children.push(PARSED_COLUMN_BREAK);
        } else if (brClear) {
          // Line break clearing floating content (w:br/@w:clear) — preserve clear
          children.push({
            break: { count: 1, clear: brClear as BreakClear },
          } as unknown as ParsedRunChild);
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
      // Drawing is handled at the paragraph level (parseSectionChild in body.ts)
      // where the drawing is extracted and replaced as a paragraph child.
      case "w:drawing":
        break;
      // VML picture — ordered shape children with imagedata media bridged
      // from the part's rels (r:id → bytes → `{fileName}` placeholder).
      case "w:pict": {
        children.push({ pict: parsePict(child, _ctx) } as unknown as ParsedRunChild);
        break;
      }
      case "w:object": {
        children.push({ object: objectDesc.parse(child, _ctx) } as unknown as ParsedRunChild);
        break;
      }
      // Symbol run — extract char and font attributes
      case "w:sym": {
        const charVal = attr(child, "w:char");
        const fontVal = attr(child, "w:font");
        if (charVal) {
          children.push({
            symbolRun: { char: charVal, symbolFont: fontVal ?? "Wingdings" },
          } as unknown as ParsedRunChild);
        }
        break;
      }
      // Footnote/endnote reference — preserve as { footnoteReference: id } / { endnoteReference: id }
      case "w:footnoteReference": {
        const id = attrNum(child, "w:id");
        if (id !== undefined) {
          const customMarkFollows = attrBool(child, "w:customMarkFollows") === true;
          children.push(
            customMarkFollows
              ? ({
                  footnoteReference: { id, customMarkFollows: true },
                } as unknown as ParsedRunChild)
              : ({ footnoteReference: id } as unknown as ParsedRunChild),
          );
        }
        break;
      }
      case "w:endnoteReference": {
        const id = attrNum(child, "w:id");
        if (id !== undefined) {
          const customMarkFollows = attrBool(child, "w:customMarkFollows") === true;
          children.push(
            customMarkFollows
              ? ({ endnoteReference: { id, customMarkFollows: true } } as unknown as ParsedRunChild)
              : ({ endnoteReference: id } as unknown as ParsedRunChild),
          );
        }
        break;
      }
      // Positional tab (EG_RunInnerContent) — absolute-positioned tab stop.
      case "w:ptab": {
        const alignment = attr(child, "w:alignment");
        const leader = attr(child, "w:leader");
        const relativeTo = attr(child, "w:relativeTo");
        if (alignment !== undefined && leader !== undefined && relativeTo !== undefined) {
          children.push({
            positionalTab: { alignment, leader, relativeTo },
          } as unknown as ParsedRunChild);
        }
        break;
      }
      // Footnote/endnote ref mark inside note content. Kept as a child so its
      // run properties round-trip (Word styles the ref mark run itself).
      case "w:footnoteRef":
        children.push(PARSED_FOOTNOTE_REF);
        break;
      case "w:endnoteRef":
        children.push(PARSED_ENDNOTE_REF);
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

  return { properties, children, additionRsid: rsid, runPropertiesRsid, deletionRsid };
}

/**
 * Convert parsed run data into an RunOptions suitable for the Document constructor.
 * Simplifies the parsed children into text + break format.
 * If the run contains only a commentReference, returns { commentReference: id } instead.
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
  [PARSED_FOOTNOTE_REF, { footnoteRef: true }],
  [PARSED_ENDNOTE_REF, { endnoteRef: true }],
  [PARSED_SEPARATOR, { separator: true }],
  [PARSED_CONTINUATION_SEPARATOR, { continuationSeparator: true }],
  [PARSED_PAGE_NUMBER, { pgNum: true }],
  [PARSED_LAST_RENDERED_PAGE_BREAK, { lastRenderedPageBreak: true }],
]);

export function parsedRunToOptions(
  parsed: ReturnType<typeof parseRun>,
): RunOptions | { commentReference: number } | null {
  const contentChildren = parsed.children;

  const opts: Record<string, unknown> = { ...parsed.properties };
  if (parsed.additionRsid) opts.additionRsid = parsed.additionRsid;
  if (parsed.runPropertiesRsid) opts.runPropertiesRsid = parsed.runPropertiesRsid;
  if (parsed.deletionRsid) opts.deletionRsid = parsed.deletionRsid;

  // Check if this run is a pure reference run (commentReference, footnoteReference, endnoteReference)
  const isRefChild = (c: unknown): c is Record<string, number> =>
    typeof c === "object" &&
    c !== null &&
    ("commentReference" in c || "footnoteReference" in c || "endnoteReference" in c);

  const refChildren = contentChildren.filter(isRefChild);
  const nonRefChildren = contentChildren.filter((c) => !isRefChild(c));
  // A reference mixed with other content (e.g. lastRenderedPageBreak before a
  // commentReference) keeps every child in children[] form — the reference
  // would otherwise be silently dropped by the text/break simplification.
  const mixedRefs = refChildren.length > 0 && nonRefChildren.length > 0;

  // If the run is a pure reference run (no text), return it directly, keeping
  // the run properties so the reference round-trips byte-faithfully.
  if (refChildren.length > 0 && nonRefChildren.length === 0) {
    const ref = refChildren[0] as { commentReference?: number };
    return Object.keys(opts).length > 0
      ? ({ ...ref, properties: opts } as RunOptions | { commentReference: number })
      : (ref as RunOptions | { commentReference: number });
  }

  // If the run only contains a symbolRun, return it directly. SymbolRunOptions
  // extends RunOptions, so the run properties merge into the symbolRun itself —
  // the stringify path serializes them from there. Without this a sym run
  // carrying w:rPr falls through to the text collection loop and is dropped.
  const symbolIdx = nonRefChildren.findIndex(
    (c) => typeof c === "object" && c !== null && "symbolRun" in c,
  );
  if (symbolIdx >= 0 && nonRefChildren.length === 1) {
    const sym = (nonRefChildren[symbolIdx] as unknown as { symbolRun: Record<string, unknown> })
      .symbolRun;
    return { symbolRun: { ...parsed.properties, ...sym } } as unknown as RunOptions;
  }

  // If the run contains only an OLE object (w:object), return it directly with
  // any run properties — an OLE object occupies its own run. Mixed with other
  // content it stays in children[] (handled below, order-preserving).
  const objectIdx = nonRefChildren.findIndex(
    (c) => typeof c === "object" && c !== null && "object" in c,
  );
  if (objectIdx >= 0 && nonRefChildren.length === 1) {
    const objectChild = nonRefChildren[objectIdx] as { object: ObjectElementOptions };
    return { ...parsed.properties, ...objectChild } as unknown as RunOptions;
  }

  // A VML picture (w:pict) likewise occupies its own run when alone.
  const pictIdx = nonRefChildren.findIndex(
    (c) => typeof c === "object" && c !== null && "pict" in c,
  );
  if (pictIdx >= 0 && nonRefChildren.length === 1) {
    const pictChild = nonRefChildren[pictIdx] as unknown as { pict: PictOptions };
    return { ...parsed.properties, ...pictChild } as unknown as RunOptions;
  }
  const hasBlockChild = objectIdx >= 0 || pictIdx >= 0;
  // A symbol run mixed with other content must stay in children[] form — the
  // text collection loop below would otherwise drop the symbolRun wrapper.
  const hasMixedSymbol = symbolIdx >= 0 && nonRefChildren.length > 1;

  // Collect text and breaks
  const textParts: string[] = [];
  let breakCount = 0;
  const structuredBreaks: BreakOptions[] = [];
  let hasPageBreak = false;
  let hasColumnBreak = false;
  const extraChildren: Record<string, unknown>[] = [];

  for (const child of nonRefChildren) {
    if (typeof child === "string") {
      textParts.push(child);
    } else if (child === PARSED_LINE_BREAK) {
      breakCount++;
    } else if (child === PARSED_PAGE_BREAK) {
      hasPageBreak = true;
    } else if (child === PARSED_COLUMN_BREAK) {
      hasColumnBreak = true;
    } else if (typeof child === "object" && child !== null && "break" in child) {
      // Line break carrying a clear attribute (w:br/@w:clear) — preserve structure
      structuredBreaks.push((child as { break: BreakOptions }).break);
    } else if (typeof child === "object" && child !== null) {
      // Valued object children the text loop cannot fold (positionalTab) —
      // keep them so useChildrenForm re-emits the wrapper as-is.
      extraChildren.push(child as Record<string, unknown>);
    } else {
      // Empty run elements (tab, noBreakHyphen, date fields, etc.)
      const mapped = SYMBOL_TO_CHILD.get(child as symbol);
      if (mapped) extraChildren.push(mapped);
    }
  }

  // A single structured break (with clear) coexists cleanly with text/page/column
  // breaks via opts.break; mixed or multiple breaks fall back to children[] form.
  const hasStructuredBreaks = structuredBreaks.length > 0;
  const useChildrenForm =
    mixedRefs ||
    (hasBlockChild && nonRefChildren.length > 1) ||
    hasMixedSymbol ||
    extraChildren.length > 0 ||
    // Multiple w:t in one run (Word splits text for session history) — joining
    // them into `text` would re-emit a single w:t and lose the split points.
    textParts.length > 1 ||
    (hasStructuredBreaks &&
      (breakCount > 0 || structuredBreaks.length > 1 || hasPageBreak || hasColumnBreak));

  if (useChildrenForm) {
    const children: (string | Record<string, unknown>)[] = [];
    for (const child of mixedRefs ? contentChildren : nonRefChildren) {
      if (typeof child === "string") {
        children.push(child);
      } else if (child === PARSED_LINE_BREAK) {
        children.push({ break: 1 });
      } else if (child === PARSED_PAGE_BREAK) {
        children.push({ pageBreak: true });
      } else if (child === PARSED_COLUMN_BREAK) {
        children.push({ columnBreak: true });
      } else if (typeof child === "object" && child !== null && "break" in child) {
        children.push({ break: (child as { break: BreakOptions }).break });
      } else if (typeof child === "object" && child !== null) {
        // Ref children, w:object/w:pict block children, symbol runs — keep the
        // parsed wrapper as-is so stringify re-emits it bare inside this run.
        children.push(child);
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
    } else if (hasStructuredBreaks) {
      opts.break = structuredBreaks[0];
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
