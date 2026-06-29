/**
 * Table of Contents parser for DOCX documents.
 *
 * Parses TOC-type SDT elements into TOC options.
 *
 * @module
 */
import { attr, children, findChild, findDeep, textOf } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import type {
  StyleLevel,
  TableOfContentsOptions,
} from "@parts/table-of-contents/table-of-contents-properties";
import type { SectionChild } from "@shared/section";

import type { DocxReadContext } from "../../context";

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

  if (sdtContent) {
    // Look for field instructions
    for (const p of children(sdtContent, "w:p")) {
      for (const r of children(p, "w:r")) {
        for (const instrText of children(r, "w:instrText")) {
          const instruction = textOf(instrText).trim();
          parseTocFieldInstruction(instruction, tocOpts);
        }
      }
    }
    // Rendered entries — the paragraphs between the field's separate and end
    // markers. Captured structurally so MS Office and WPS both display the
    // existing TOC instead of regenerating it from headings.
    if (parseChildren) {
      const entryEls = selectTocEntryElements(sdtContent.elements ?? []);
      if (entryEls.length > 0) {
        tocOpts.entries = parseChildren(entryEls, ctx);
      }
    }
  }

  return { alias, ...tocOpts } as { alias?: string } & TableOfContentsOptions;
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
      const level = parseInt(parts[i + 1], 10);
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
 * a bare cross-paragraph field). Feeds every w:instrText to the instruction
 * parser; non-TOC fields (HYPERLINK/PAGEREF inside the rendered entries) are
 * ignored — parseTocFieldInstruction only acts on instructions starting "TOC".
 */
export function parseTocFieldFromElements(els: Element[]): TableOfContentsOptions {
  const opts: Record<string, unknown> = {};
  for (const el of els) collectTocInstructions(el, opts);
  return opts as TableOfContentsOptions;
}

/**
 * Select the rendered-entry paragraphs of a captured TOC field. Tracks field
 * depth so a nested HYPERLINK/PAGEREF field inside an entry doesn't fool the
 * boundary detection.
 *
 * A paragraph is an entry when, after walking it, the field is past `separate`
 * and not past `end` (depth ≥ 1) and the paragraph carries rendered text (`w:t`).
 * This captures an entry whose paragraph also opens the field (`begin`) or holds
 * the `separate` marker — common when Word emits begin + separate + first entry
 * in one paragraph — while the `w:t` requirement excludes a pure control
 * paragraph (field head / separate-only / end).
 */
export function selectTocEntryElements(els: Element[]): Element[] {
  const entries: Element[] = [];
  let depth = 0;
  let afterSeparate = false;
  for (const el of els) {
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
    if (afterSeparate && depth >= 1 && findDeep(el, "w:t").length > 0) {
      entries.push(el);
    }
  }
  return entries;
}

/** Recursively feed every w:instrText to the TOC instruction parser. */
function collectTocInstructions(el: Element, opts: Record<string, unknown>): void {
  if (el.name === "w:instrText") {
    const instruction = textOf(el)?.trim();
    if (instruction) parseTocFieldInstruction(instruction, opts);
  }
  for (const c of el.elements ?? []) {
    if (c.type === "element") collectTocInstructions(c, opts);
  }
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
    while (i < text.length && /[a-zA-Z]/.test(text[i])) i++;
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
