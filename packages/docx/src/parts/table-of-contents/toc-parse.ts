/**
 * Table of Contents parser for DOCX documents.
 *
 * Parses TOC-type SDT elements into TOC options.
 *
 * @module
 */
import { attr, children, findChild, textOf } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import type { TableOfContentsOptions } from "@parts/table-of-contents/table-of-contents-properties";

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
  _ctx: DocxReadContext,
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
function parseTocFieldInstruction(instruction: string, opts: Record<string, unknown>): void {
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
  if ("u" in switches) opts.useAppliedParagraphOutlineLevel = true;
  if ("w" in switches) opts.preserveTabInEntries = true;
  if ("x" in switches) opts.preserveNewLineInEntries = true;
  if ("z" in switches) opts.hideTabAndPageNumbersInWebView = true;
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
