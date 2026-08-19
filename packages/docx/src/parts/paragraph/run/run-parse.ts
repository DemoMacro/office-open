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
const ON_OFF_RUN_PROPS_MAP: ReadonlyMap<string, string> = new Map(ON_OFF_RUN_PROPS);

/**
 * Parse a w:rPr element into RunPropertiesOptions.
 *
 * Single pass over the children: rPr holds at most one of each property, so a
 * switch dispatch beats the previous per-property findChild linear scans
 * (26 properties × N children re-walked the array for every lookup).
 */
export function parseRunProperties(el: Element): RunPropertiesOptions {
  const opts: Record<string, unknown> = {};
  let w14Parts: string[] | undefined;

  const children = el.elements;
  if (children !== undefined) {
    for (let i = 0; i < children.length; i++) {
      const child = children[i];
      if (child === undefined || child.type !== "element") continue;
      switch (child.name) {
        case "w:rStyle":
          opts.style = attr(child, "w:val");
          break;
        case "w:rFonts": {
          const ascii = attr(child, "w:ascii");
          const eastAsia = attr(child, "w:eastAsia");
          const hAnsi = attr(child, "w:hAnsi");
          const complexScript = attr(child, "w:cs");
          const asciiTheme = attr(child, "w:asciiTheme");
          const eastAsiaTheme = attr(child, "w:eastAsiaTheme");
          const hAnsiTheme = attr(child, "w:hAnsiTheme");
          const cstheme = attr(child, "w:cstheme");
          const hint = attr(child, "w:hint");

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
          break;
        }
        case "w:b":
          opts.bold = attrBool(child, "w:val") ?? true;
          break;
        case "w:bCs":
          opts.boldComplexScript = attrBool(child, "w:val") ?? true;
          break;
        case "w:i":
          opts.italic = attrBool(child, "w:val") ?? true;
          break;
        case "w:iCs":
          opts.italicComplexScript = attrBool(child, "w:val") ?? true;
          break;
        case "w:u": {
          const ul: Record<string, string | undefined> = {};
          const uType = attr(child, "w:val");
          if (uType) ul.type = uType;
          const uColor = colorAttr(child, "w:color");
          if (uColor) ul.color = uColor;
          opts.underline = ul;
          break;
        }
        case "w:color": {
          const c = colorAttr(child, "w:val");
          const themeColor = attr(child, "w:themeColor");
          const themeTint = attr(child, "w:themeTint");
          const themeShade = attr(child, "w:themeShade");
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
          break;
        }
        case "w:sz": {
          const halfPts = attrNum(child, "w:val");
          if (halfPts !== undefined) opts.size = halfPts / 2;
          break;
        }
        case "w:szCs": {
          const halfPts = attrNum(child, "w:val");
          if (halfPts !== undefined) opts.sizeComplexScript = halfPts / 2;
          break;
        }
        case "w:highlight": {
          const val = attr(child, "w:val");
          if (val) opts.highlight = val;
          break;
        }
        case "w:vertAlign": {
          const val = attr(child, "w:val");
          if (val === "baseline" || val === "subscript" || val === "superscript") {
            opts.verticalAlign = val;
          }
          break;
        }
        case "w:effect": {
          const val = attr(child, "w:val");
          if (val) opts.effect = val;
          break;
        }
        case "w:em": {
          const val = attr(child, "w:val");
          if (val) opts.emphasisMark = { type: val };
          break;
        }
        case "w:spacing": {
          const val = attrMeasure(child, "w:val");
          if (val !== undefined) opts.characterSpacing = val;
          break;
        }
        case "w:w": {
          const val = attrNum(child, "w:val");
          if (val !== undefined) opts.scale = val;
          break;
        }
        case "w:kern": {
          // w:kern is ST_HpsMeasure: numeric tokens are half-points → points
          // (÷2), UniversalMeasure strings pass through verbatim (same split
          // as w:position).
          const val = attrMeasure(child, "w:val");
          if (val !== undefined) {
            opts.kern = typeof val === "number" ? val / 2 : (val as UniversalMeasure);
          }
          break;
        }
        case "w:position": {
          // w:position is ST_SignedHpsMeasure; numeric tokens are half-points
          // → points (÷2), UniversalMeasure strings pass through verbatim.
          const val = attrMeasure(child, "w:val");
          if (val !== undefined) {
            opts.position = typeof val === "number" ? val / 2 : (val as UniversalMeasure);
          }
          break;
        }
        case "w:fitText": {
          const val = attrNum(child, "w:val");
          if (val !== undefined) opts.fitText = val;
          break;
        }
        case "w:lang": {
          // Keep the element even when it carries no attributes — Word writes
          // a bare <w:lang/> to override inherited language settings.
          const langObj: LanguageOptions = {};
          const val = attr(child, "w:val");
          if (val) langObj.value = val;
          const eastAsia = attr(child, "w:eastAsia");
          if (eastAsia) langObj.eastAsia = eastAsia;
          const bidi = attr(child, "w:bidi");
          if (bidi) langObj.bidirectional = bidi;
          opts.language = langObj;
          break;
        }
        case "w:bdr":
          opts.border = parseBorder(child);
          break;
        case "w:shd":
          opts.shading = parseShading(child);
          break;
        case "w:eastAsianLayout":
          opts.eastAsianLayout = parseEastAsianLayout(child);
          break;
        case "w:contentPart": {
          const rId = attr(child, "r:id");
          if (rId) opts.contentPartRId = rId;
          break;
        }
        case "w:rPrChange": {
          const rev: Record<string, unknown> = {};
          const author = attr(child, "w:author");
          if (author) rev.author = author;
          const date = attr(child, "w:date");
          if (date) rev.date = date;
          const id = attrNum(child, "w:id");
          if (id !== undefined) rev.id = id;
          const innerRPr = findChild(child, "w:rPr");
          if (innerRPr) {
            Object.assign(rev, parseRunProperties(innerRPr));
          }
          if (Object.keys(rev).length > 0) opts.revision = rev;
          break;
        }
        default: {
          // On/off properties (w:strike, w:caps, …) and w14:* text effects.
          // w14 effects (glow/shadow/reflection/props3d) occupy the
          // EG_RPrBase extension slot — low-frequency complex subtrees kept
          // verbatim as raw XML for fidelity while the rPr backbone stays
          // editable.
          const name = child.name;
          if (name === undefined) break;
          const optKey = ON_OFF_RUN_PROPS_MAP.get(name);
          if (optKey !== undefined) {
            opts[optKey] = attrBool(child, "w:val") ?? true;
          } else if (name.startsWith("w14:")) {
            (w14Parts ??= []).push(stringifyElement(child));
          }
          break;
        }
      }
    }
  }

  if (w14Parts !== undefined && w14Parts.length > 0) opts.w14RawXml = w14Parts.join("");

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

  // Fast path: the overwhelmingly common run shape — one plain text node, no
  // rsids — skips the reference/symbol/object scans and collection loop.
  if (
    contentChildren.length === 1 &&
    typeof contentChildren[0] === "string" &&
    parsed.additionRsid === undefined &&
    parsed.runPropertiesRsid === undefined &&
    parsed.deletionRsid === undefined
  ) {
    const text = contentChildren[0];
    return parsed.properties === undefined
      ? ({ text } as RunOptions)
      : ({ ...parsed.properties, text } as RunOptions);
  }

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
