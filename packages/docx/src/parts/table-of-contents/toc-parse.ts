/**
 * Table of Contents parser for DOCX documents.
 *
 * Parses TOC-type SDT elements into TOC options.
 *
 * @module
 */
import { attr, children, findChild, findFirst, textOf } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { parseRunProperties } from "@parts/paragraph/run/run-parse";
import type {
  StyleLevel,
  TableOfContentsOptions,
} from "@parts/table-of-contents/table-of-contents-properties";
import type { SectionChild } from "@shared/section";

import type { DocxReadContext } from "../../context";
import { captureTocFieldRPr } from "../../parse/body";

/**
 * Try to parse a w:sdt element as a TOC.
 * Returns { alias, ...tocOptions } if it's a TOC, or undefined otherwise.
 *
 * Detects TOC in two ways:
 * 1. SDT with w:docPartObj > w:docPartGallery = "Table of Contents" (Word-generated)
 * 2. SDT whose content contains a TOC field instruction (library-generated)
 */
export function parseToc(
  el: Element,
  ctx: DocxReadContext,
  parseChildren?: (elements: Element[], ctx: DocxReadContext) => SectionChild[],
):
  | ({
      alias?: string;
    } & TableOfContentsOptions)
  | undefined {
  const sdtPr = findChild(el, "w:sdtPr");
  if (!sdtPr) return undefined;

  // Detection method 1: docPartObj with gallery "Table of Contents"
  const docPartObj = findChild(sdtPr, "w:docPartObj");
  let isToc = false;

  if (docPartObj) {
    const gallery = findChild(docPartObj, "w:docPartGallery");
    if (gallery && textOf(gallery) === "Table of Contents") {
      isToc = true;
    }
  }

  // Detection method 2: scan content for TOC field instruction
  if (!isToc) {
    const sdtContent = findChild(el, "w:sdtContent");
    if (sdtContent) {
      isToc = hasTocFieldInstruction(sdtContent);
    }
  }

  if (!isToc) return undefined;

  // It's a TOC SDT
  const alias = (() => {
    const aliasEl = findChild(sdtPr, "w:alias");
    return aliasEl ? attr(aliasEl, "w:val") : undefined;
  })();

  // Parse field instruction from content to extract TOC options
  const tocOpts: Record<string, unknown> = {};
  const sdtContent = findChild(el, "w:sdtContent");

  // sdtPr round-trip fields: the start mark's rPr, the SDT id, and
  // docPartObj's w:docPartUnique (gallery is the TOC detection key).
  const rPr = findChild(sdtPr, "w:rPr");
  if (rPr) tocOpts.runProperties = parseRunProperties(rPr);
  const idEl = findChild(sdtPr, "w:id");
  if (idEl) {
    const val = attr(idEl, "w:val");
    const num = val !== undefined ? Number(val) : NaN;
    if (!Number.isNaN(num)) tocOpts.id = num;
  }
  if (docPartObj && findChild(docPartObj, "w:docPartUnique")) tocOpts.docPartUnique = true;

  // CT_SdtEndPr wraps its run properties in a w:rPr child; a bare element is
  // carried as an empty object so stringify re-emits it.
  const sdtEndPr = findChild(el, "w:sdtEndPr");
  if (sdtEndPr) {
    const endRPr = findChild(sdtEndPr, "w:rPr");
    tocOpts.endProperties = endRPr ? parseRunProperties(endRPr) : {};
  }

  if (sdtContent) {
    const elements = sdtContent.elements ?? [];
    // Word splits one instruction across runs (` TOC ` + ` \o "1-1" `), so
    // the outer field's instrText runs are joined before parsing.
    const instruction = collectOuterFieldInstruction(elements);
    if (instruction) parseTocFieldInstruction(instruction, tocOpts);
    // Capture the field's control runs verbatim (begin→separate chain, the
    // control rPr, the end run's rPr) — same contract as the bare-field TOC.
    captureTocFieldRPr(elements, tocOpts as TableOfContentsOptions);
    // sdtContent may carry block content around the field — the TOC heading
    // paragraph before it, the wrapping bookmarkStart/bookmarkEnd pair. Split
    // the element list at the field's outer begin/end so nothing drops.
    const { leadingEls, fieldEls, trailingEls } = splitAroundOuterField(elements);
    if (parseChildren) {
      if (leadingEls.length > 0) tocOpts.leading = parseChildren(leadingEls, ctx);
      if (trailingEls.length > 0) tocOpts.trailing = parseChildren(trailingEls, ctx);
      // Rendered entries — the paragraphs between the field's separate and end
      // markers. Captured structurally so MS Office and WPS both display the
      // existing TOC instead of regenerating it from headings.
      const entryEls = selectTocEntryElements(fieldEls);
      if (entryEls.length > 0) {
        tocOpts.entries = parseChildren(entryEls, ctx);
      }
    }
  }

  return { alias, ...tocOpts } as { alias?: string } & TableOfContentsOptions;
}

/**
 * Split sdtContent's element list into leading / field / trailing around the
 * outermost complex field. The field span runs from the element that opens
 * the outer field (depth 0→1) through the element that closes it (back to 0);
 * an unterminated field (its end marker lives in a later body paragraph —
 * endInBody) keeps the remainder in the field slice.
 */
function splitAroundOuterField(els: Element[]): {
  leadingEls: Element[];
  fieldEls: Element[];
  trailingEls: Element[];
} {
  let depth = 0;
  let startIdx = -1;
  let endIdx = -1;
  for (let i = 0; i < els.length; i++) {
    const depthBefore = depth;
    let opened = false;
    const walk = (node: Element): void => {
      if (node.name === "w:fldChar") {
        const type = attr(node, "w:fldCharType");
        if (type === "begin") {
          depth++;
          opened = true;
        } else if (type === "end") depth--;
      }
      for (const c of node.elements ?? []) {
        if (c.type === "element") walk(c);
      }
    };
    walk(els[i]!);
    // `opened` (not depth > depthBefore): an element may open AND close the
    // field itself — a single-entry TOC renders the whole field in one
    // paragraph — in which case depth returns to depthBefore and the element
    // must still start the field slice.
    if (startIdx < 0 && opened) startIdx = i;
    // The field span closes when depth returns to 0 from an open field, or in
    // the element that opened it (the single-paragraph case), so content after
    // the field stays in the trailing slice instead of being swallowed.
    if (startIdx >= 0 && endIdx < 0 && depth === 0 && (depthBefore > 0 || opened)) {
      endIdx = i;
      break;
    }
  }
  if (startIdx < 0) return { leadingEls: els, fieldEls: [], trailingEls: [] };
  return {
    leadingEls: els.slice(0, startIdx),
    fieldEls: els.slice(startIdx, endIdx >= 0 ? endIdx + 1 : els.length),
    trailingEls: endIdx >= 0 ? els.slice(endIdx + 1) : [],
  };
}

/**
 * Check whether an element tree contains a TOC field instruction
 * (w:instrText with text starting with "TOC").
 */
function hasTocFieldInstruction(el: Element): boolean {
  for (const p of children(el, "w:p")) {
    for (const r of children(p, "w:r")) {
      for (const instrText of children(r, "w:instrText")) {
        const instruction = textOf(instrText)?.trim();
        if (instruction?.startsWith("TOC")) return true;
      }
    }
  }
  return false;
}

/**
 * Parse a TOC field instruction string (e.g., ' TOC \o "1-3" \h \z ')
 * into TableOfContentsOptions properties.
 */
export function parseTocFieldInstruction(instruction: string, opts: Record<string, unknown>): void {
  if (!instruction.startsWith("TOC")) return;

  const rest = instruction.slice(3).trim();
  const switches = parseFieldSwitches(rest);

  if (switches["a"]) opts.captionLabel = switches["a"];
  if (switches["b"]) opts.entriesFromBookmark = switches["b"];
  if (switches["c"]) opts.captionLabelIncludingNumbers = switches["c"];
  if (switches["d"]) opts.sequenceAndPageNumbersSeparator = switches["d"];
  if (switches["f"]) opts.tcFieldIdentifier = switches["f"];
  if ("h" in switches) opts.hyperlink = true;
  if (switches["l"]) opts.tcFieldLevelRange = switches["l"];
  if (switches["n"]) opts.pageNumbersEntryLevelsRange = switches["n"];
  if (switches["o"]) opts.headingStyleRange = switches["o"];
  if (switches["p"]) opts.entryAndPageNumberSeparator = switches["p"];
  if (switches["s"]) opts.seqFieldIdentifierForPrefix = switches["s"];
  if (switches["t"]) {
    // \t "Style1,1,Style2,2" -> stylesWithLevels pairs
    const parts = switches["t"]!.split(",");
    const stylesWithLevels: StyleLevel[] = [];
    for (let i = 0; i + 1 < parts.length; i += 2) {
      const styleName = parts[i];
      const level = parseInt(parts[i + 1] ?? "", 10);
      if (styleName && !Number.isNaN(level)) stylesWithLevels.push({ styleName, level });
    }
    if (stylesWithLevels.length > 0) opts.stylesWithLevels = stylesWithLevels;
  }
  if ("u" in switches) opts.useAppliedParagraphOutlineLevel = true;
  if ("w" in switches) opts.preserveTabInEntries = true;
  if ("x" in switches) opts.preserveNewLineInEntries = true;
  if ("z" in switches) opts.hideTabAndPageNumbersInWebView = true;
}

/**
 * Extract TOC options from the elements of a captured TOC field (SDT content or
 * a bare cross-paragraph field). Non-TOC fields (HYPERLINK/PAGEREF inside the
 * rendered entries) never reach the instruction parser — see
 * collectOuterFieldInstruction.
 */
export function parseTocFieldFromElements(els: Element[]): TableOfContentsOptions {
  const opts: Record<string, unknown> = {};
  const instruction = collectOuterFieldInstruction(els);
  if (instruction) parseTocFieldInstruction(instruction, opts);
  return opts as TableOfContentsOptions;
}

/**
 * Select the rendered-entry paragraphs of a captured TOC field. Tracks field
 * depth so a nested HYPERLINK/PAGEREF field inside an entry doesn't fool the
 * boundary detection.
 *
 * A paragraph is an entry when, after walking it, the field is past `separate`
 * and the paragraph carries rendered text (`w:t`). This captures an entry whose
 * paragraph also opens the field (`begin`) or holds the `separate` marker or
 * closes the field in the same paragraph — Word may render the whole field in
 * one paragraph — while the `w:t` requirement excludes a pure control
 * paragraph (field head / separate-only / end).
 *
 * A paragraph that closes an already-open field (depth drops to 0) joins the
 * entries when it is a pure control paragraph the source markup carried or
 * when it is itself the last rendered entry (see keepTocClosingParagraph);
 * a text-bearing paragraph the end merely drifted into stays body content.
 */
export function selectTocEntryElements(els: Element[]): Element[] {
  const entries: Element[] = [];
  let depth = 0;
  let afterSeparate = false;
  for (const el of els) {
    const depthBefore = depth;
    const walk = (node: Element): void => {
      if (node.name === "w:fldChar") {
        const type = attr(node, "w:fldCharType");
        if (type === "begin") depth++;
        else if (type === "separate" && depth === 1) afterSeparate = true;
        else if (type === "end") depth--;
      }
      for (const c of node.elements ?? []) {
        if (c.type === "element") walk(c);
      }
    };
    walk(el);
    if (afterSeparate && depth === 0) {
      // depthBefore === 0 marks a paragraph that opens AND closes the field in
      // itself (single-paragraph TOC): it carries the entry text, not a
      // trailing control paragraph, so plain text membership decides.
      const closing = depthBefore > 0;
      const keep = closing
        ? keepTocClosingParagraph(el, entries.length > 0)
        : findFirst(el, "w:t") !== undefined;
      if (keep) entries.push(el);
      afterSeparate = false;
    } else if (afterSeparate && depth >= 1 && findFirst(el, "w:t") !== undefined) {
      entries.push(el);
    }
  }
  return entries;
}

/**
 * Whether the field-closing paragraph joins the rendered entries: a pure
 * control paragraph (no rendered text) that either carries a pPr (Word parks
 * the entry style's properties there) or follows at least one rendered entry
 * — a paragraph the source markup carried, so the stringify path injects the
 * field-end run back into it, keeping the paragraph count identical. A bare
 * closing paragraph with no entries before it belongs to a never-rendered
 * field (fresh dirty TOC) and is dropped. A text-bearing closing paragraph
 * joins the entries only when it is itself the last rendered entry (see
 * isTocEntryClosingParagraph); a paragraph the end merely drifted into is
 * body content and must stay in the body, not the TOC.
 */
export function keepTocClosingParagraph(el: Element, hasEntries: boolean): boolean {
  if (findFirst(el, "w:t") === undefined) {
    return hasEntries || findFirst(el, "w:pPr") !== undefined;
  }
  return isTocEntryClosingParagraph(el);
}

/**
 * True when a text-bearing paragraph that closes the TOC field is itself the
 * last rendered entry. The writer parks the field end in the last entry's
 * paragraph, so such a paragraph keeps an entry's markup — a TOC style, a
 * hyperlink, or a nested page-number field — and must stay in the entries
 * rather than being dropped/moved to the body. A paragraph the end merely
 * drifted into (a following heading) carries none of those.
 */
function isTocEntryClosingParagraph(el: Element): boolean {
  const pPr = findChild(el, "w:pPr");
  const styleEl = pPr ? findChild(pPr, "w:pStyle") : undefined;
  const style = styleEl ? attr(styleEl, "w:val") : undefined;
  if (typeof style === "string" && style.toUpperCase().startsWith("TOC")) return true;
  if (findFirst(el, "w:hyperlink") !== undefined) return true;
  let fields = 0;
  const walk = (node: Element): void => {
    if (node.name === "w:fldChar") fields++;
    for (const c of node.elements ?? []) {
      if (c.type === "element") walk(c);
    }
  };
  walk(el);
  return fields > 1;
}

/**
 * Concatenate the instrText runs of the OUTERMOST field in the given
 * elements, trimmed. Word splits one instruction across runs (` TOC ` +
 * ` \o "1-1" `), so the runs must be joined before parsing — feeding each
 * run on its own loses every switch that lands in a later run. Nested
 * fields (HYPERLINK/PAGEREF inside the rendered entries) sit at depth ≥ 2
 * and are excluded, so their switches cannot leak into the TOC's.
 */
function collectOuterFieldInstruction(els: Element[]): string {
  let depth = 0;
  let inHead = false;
  let instruction = "";
  const walk = (node: Element): void => {
    if (node.name === "w:r" && node.elements) {
      // A run holds at most one fldChar marker; instrText runs carry none.
      // Handle begin → instrText → separate in document order so a run
      // combining markers still opens/closes the head window correctly.
      for (const c of node.elements) {
        if (c.type !== "element") continue;
        if (c.name === "w:fldChar") {
          const type = attr(c, "w:fldCharType");
          if (type === "begin") {
            depth++;
            if (depth === 1) inHead = true;
          } else if (type === "separate") {
            if (depth === 1) inHead = false;
          } else if (type === "end") {
            if (depth === 1) inHead = false;
            depth--;
          }
        } else if (c.name === "w:instrText" && inHead && depth === 1) {
          instruction += textOf(c) ?? "";
        }
      }
      return;
    }
    for (const c of node.elements ?? []) {
      if (c.type === "element") walk(c);
    }
  };
  for (const el of els) walk(el);
  return instruction.trim();
}

/**
 * Parse field switches like \o "1-3" \h \z into a map.
 */
function parseFieldSwitches(text: string): Record<string, string | undefined> {
  const result: Record<string, string | undefined> = {};
  let i = 0;

  while (i < text.length) {
    // Skip whitespace
    while (i < text.length && text[i] === " ") i++;
    if (i >= text.length) break;

    // Expect backslash
    if (text[i] !== "\\") {
      i++;
      continue;
    }
    i++;

    // Read switch name
    const nameStart = i;
    while (i < text.length && /[a-zA-Z]/.test(text[i] ?? "")) i++;
    const name = text.slice(nameStart, i);
    if (!name) continue;

    // Skip whitespace
    while (i < text.length && text[i] === " ") i++;

    // Read argument (quoted or unquoted)
    if (i < text.length && text[i] === '"') {
      i++;
      const argStart = i;
      while (i < text.length && text[i] !== '"') i++;
      result[name] = text.slice(argStart, i);
      if (i < text.length) i++; // skip closing quote
    } else if (i < text.length && text[i] !== "\\") {
      const argStart = i;
      while (i < text.length && text[i] !== " " && text[i] !== "\\") i++;
      const arg = text.slice(argStart, i).trim();
      if (arg) result[name] = arg;
      else result[name] = undefined;
    } else {
      result[name] = undefined;
    }
  }

  return result;
}
