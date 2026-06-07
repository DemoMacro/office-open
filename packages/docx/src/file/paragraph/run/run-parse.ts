import type { RunPropertiesOptions, RunOptions } from "@file/paragraph/run";
import { RawPassthrough } from "@office-open/core";
/**
 * Run properties parser for DOCX documents.
 *
 * Parses w:rPr Element trees into RunPropertiesOptions objects.
 *
 * @module
 */
import { attr, attrBool, attrNum, colorAttr, findChild, textOf } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { ParseContext } from "../../../parse/context";

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
    const hint = attr(font, "w:hint");

    if (ascii && !eastAsia && !hAnsi && !cs) {
      opts.font = hint ? { name: ascii, hint } : ascii;
    } else {
      const fontObj: Record<string, string | undefined> = {};
      if (ascii) fontObj.ascii = ascii;
      if (eastAsia) fontObj.eastAsia = eastAsia;
      if (hAnsi) fontObj.hAnsi = hAnsi;
      if (cs) fontObj.cs = cs;
      if (hint) fontObj.hint = hint;
      opts.font = fontObj;
    }
  }

  const bold = findChild(el, "w:b");
  if (bold) opts.bold = attrBool(bold, "w:val") ?? true;

  const boldCs = findChild(el, "w:bCs");
  if (boldCs) opts.boldComplexScript = attrBool(boldCs, "w:val") ?? true;

  const italics = findChild(el, "w:i");
  if (italics) opts.italics = attrBool(italics, "w:val") ?? true;

  const italicsCs = findChild(el, "w:iCs");
  if (italicsCs) opts.italicsComplexScript = attrBool(italicsCs, "w:val") ?? true;

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
    ["w:math", "math"],
  ] as const) {
    const child = findChild(el, name);
    if (child) opts[optKey] = attrBool(child, "w:val") ?? true;
  }

  const color = findChild(el, "w:color");
  if (color) {
    const c = colorAttr(color, "w:val");
    if (c) opts.color = c;
  }

  const sz = findChild(el, "w:sz");
  if (sz) {
    const halfPts = attrNum(sz, "w:val");
    if (halfPts !== undefined) opts.size = halfPts;
  }

  const szCs = findChild(el, "w:szCs");
  if (szCs) {
    const halfPts = attrNum(szCs, "w:val");
    if (halfPts !== undefined) opts.sizeComplexScript = halfPts;
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
    const val = attrNum(spacing, "w:val");
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
    const langObj: Record<string, string | undefined> = {};
    const val = attr(lang, "w:val");
    if (val) langObj.val = val;
    const eastAsia = attr(lang, "w:eastAsia");
    if (eastAsia) langObj.eastAsia = eastAsia;
    const bidi = attr(lang, "w:bidi");
    if (bidi) langObj.bidi = bidi;
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

// ── Special run child constants ──────────────────────────────────────────────

/** Matches w:br[@w:type="page"] → PageBreak */
export const PARSED_PAGE_BREAK = Symbol("PageBreak");
/** Matches w:br (line break) */
export const PARSED_LINE_BREAK = Symbol("LineBreak");
/** Matches w:tab */
export const PARSED_TAB = Symbol("Tab");
/** Matches w:cr */
export const PARSED_CR = Symbol("CarriageReturn");
/** Matches w:noBreakHyphen */
export const PARSED_NO_BREAK_HYPHEN = Symbol("NoBreakHyphen");
/** Matches w:softHyphen */
export const PARSED_SOFT_HYPHEN = Symbol("SoftHyphen");
/** Matches w:footnoteRef — auto-generated by Footnote class */
export const PARSED_FOOTNOTE_REF = Symbol("FootnoteRef");
/** Matches w:br[@w:type="column"] */
export const PARSED_COLUMN_BREAK = Symbol("ColumnBreak");

export type ParsedRunChild =
  | string
  | typeof PARSED_PAGE_BREAK
  | typeof PARSED_LINE_BREAK
  | typeof PARSED_TAB
  | typeof PARSED_CR
  | typeof PARSED_NO_BREAK_HYPHEN
  | typeof PARSED_SOFT_HYPHEN
  | typeof PARSED_FOOTNOTE_REF
  | typeof PARSED_COLUMN_BREAK
  | { commentReference: number }
  | RawPassthrough;

/**
 * Parse a w:r element into run data.
 * Returns { properties, children } where children are parsed run content items.
 */
export function parseRun(
  el: Element,
  _ctx: ParseContext,
): {
  properties: RunPropertiesOptions | undefined;
  children: ParsedRunChild[];
} {
  const rPr = findChild(el, "w:rPr");
  const properties = rPr ? parseRunProperties(rPr) : undefined;
  const children: ParsedRunChild[] = [];

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
        children.push(PARSED_CR);
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
      default:
        // Preserve unknown run-level elements for round-trip fidelity
        if (child.name && child.elements && child.elements.length > 0) {
          children.push(new RawPassthrough(child));
        }
        break;
    }
  }

  return { properties, children };
}

/**
 * Convert parsed run data into an RunOptions suitable for the Document constructor.
 * Simplifies the parsed children into text + break format.
 * If the run contains only a commentReference, returns { commentReference: id } instead.
 * If the run only contains footnoteRef/endnoteRef (auto-generated), returns empty options.
 */
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

  for (const child of nonRefChildren) {
    if (typeof child === "string") {
      textParts.push(child);
    } else if (child === PARSED_LINE_BREAK) {
      breakCount++;
    } else if (child === PARSED_PAGE_BREAK) {
      hasPageBreak = true;
    } else if (child === PARSED_COLUMN_BREAK) {
      hasColumnBreak = true;
    }
    // Other symbols (tab, etc.) are dropped in simplified mode
  }

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

  // If the run has no content and no properties (e.g., a pure drawing run),
  // return null so it can be skipped by the caller.
  if (
    Object.keys(opts).length === 0 &&
    textParts.length === 0 &&
    breakCount === 0 &&
    !hasPageBreak &&
    !hasColumnBreak
  ) {
    return null;
  }

  return opts as RunOptions;
}
