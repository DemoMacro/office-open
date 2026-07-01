/**
 * Table of Contents stringifier for DOCX.
 *
 * Produces `<w:sdt>` XML containing a TOC field code, replacing
 * `new TableOfContents(alias, options).toXml(ctx)` with direct
 * string concatenation — zero XmlComponent instances.
 *
 * @module
 */

import { escapeXml } from "@office-open/xml";
import type { TableOfContentsOptions } from "@parts/table-of-contents/table-of-contents-properties";

// ── Field instruction string ──

function tocInstructionStr(opts: TableOfContentsOptions): string {
  let instr = "TOC";

  if (opts.captionLabel) instr += ` \\a "${opts.captionLabel}"`;
  if (opts.entriesFromBookmark) instr += ` \\b "${opts.entriesFromBookmark}"`;
  if (opts.captionLabelIncludingNumbers) instr += ` \\c "${opts.captionLabelIncludingNumbers}"`;
  if (opts.sequenceAndPageNumbersSeparator)
    instr += ` \\d "${opts.sequenceAndPageNumbersSeparator}"`;
  if (opts.tcFieldIdentifier) instr += ` \\f "${opts.tcFieldIdentifier}"`;
  if (opts.hyperlink) instr += " \\h";
  if (opts.tcFieldLevelRange) instr += ` \\l "${opts.tcFieldLevelRange}"`;
  if (opts.pageNumbersEntryLevelsRange) instr += ` \\n "${opts.pageNumbersEntryLevelsRange}"`;
  if (opts.headingStyleRange) instr += ` \\o "${opts.headingStyleRange}"`;
  if (opts.entryAndPageNumberSeparator) instr += ` \\p "${opts.entryAndPageNumberSeparator}"`;
  if (opts.seqFieldIdentifierForPrefix) instr += ` \\s "${opts.seqFieldIdentifierForPrefix}"`;
  if (opts.stylesWithLevels?.length) {
    const styles = opts.stylesWithLevels.map((sl) => `${sl.styleName},${sl.level}`).join(",");
    instr += ` \\t "${styles}"`;
  }
  if (opts.useAppliedParagraphOutlineLevel) instr += " \\u";
  if (opts.preserveTabInEntries) instr += " \\w";
  if (opts.preserveNewLineInEntries) instr += " \\x";
  if (opts.hideTabAndPageNumbersInWebView) instr += " \\z";

  return instr;
}

// ── Main stringifier ──

export function stringifyTableOfContents(
  alias: string = "Table of Contents",
  options: TableOfContentsOptions = {},
  entriesXml: string = "",
): string {
  const instr = tocInstructionStr(options);
  const aliasAttr = alias ? ` w:val="${escapeXml(alias)}"` : "";

  // When the rendered entries are carried (round-trip), emit the field clean so
  // both MS Office and WPS display the existing TOC without an update prompt.
  // A freshly generated TOC carries no entries — mark it dirty so the consuming
  // application builds them from headings on open.
  const dirtyAttr = entriesXml.length > 0 ? "" : ' w:dirty="1"';

  // Word shares the field-head control runs (begin/instr/separate) with the
  // first rendered entry's paragraph and the field-end run with the last,
  // rather than emitting standalone control-only paragraphs — those have no
  // w:t and render as blank lines above and below the TOC. So when entries are
  // carried we inject head into the first entry paragraph and end into the
  // last; a freshly generated TOC (no entries) keeps standalone head/end
  // paragraphs since it is dirty and rebuilt on open.
  const headRuns =
    `<w:r><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:cstheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cs="Times New Roman"/></w:rPr><w:fldChar w:fldCharType="begin"${dirtyAttr}/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>`;
  const endRun = `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
  const endParagraph = `<w:p>${endRun}</w:p>`;

  const body = entriesXml
    ? injectFieldEnd(injectFieldHead(entriesXml, headRuns), endRun)
    : `<w:p>${headRuns}</w:p>` + endParagraph;

  // SDT properties: alias + docPartObj
  const sdtPr =
    `<w:sdtPr>` +
    `<w:alias${aliasAttr}/>` +
    `<w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj>` +
    `</w:sdtPr>`;

  const content = `<w:sdtContent>${body}</w:sdtContent>`;

  return `<w:sdt>${sdtPr}${content}</w:sdt>`;
}

/**
 * Inject the field-head runs into the first `<w:p>` of `entriesXml`, placing
 * them after the opening tag (and after `<w:pPr>` when present) so the head
 * shares the first entry's paragraph. Returns `entriesXml` unchanged when no
 * `<w:p>` is found.
 */
function injectFieldHead(entriesXml: string, headRuns: string): string {
  const pTagStart = entriesXml.search(/<w:p[ >]/);
  if (pTagStart < 0) return entriesXml;
  const pTagEnd = entriesXml.indexOf(">", pTagStart) + 1;
  let injectAt = pTagEnd;
  if (entriesXml.slice(pTagEnd, pTagEnd + 7) === "<w:pPr>") {
    const pPrEnd = entriesXml.indexOf("</w:pPr>", pTagEnd);
    if (pPrEnd >= 0) injectAt = pPrEnd + "</w:pPr>".length;
  }
  return entriesXml.slice(0, injectAt) + headRuns + entriesXml.slice(injectAt);
}

/**
 * Inject the field-end run into the last `<w:p>` of `entriesXml` (before its
 * closing `</w:p>`) so the end shares the last entry's paragraph instead of
 * occupying a standalone control-only paragraph that renders as a blank line.
 * Returns `entriesXml` unchanged when no `</w:p>` is found.
 */
function injectFieldEnd(entriesXml: string, endRun: string): string {
  const lastClose = entriesXml.lastIndexOf("</w:p>");
  if (lastClose < 0) return entriesXml;
  return entriesXml.slice(0, lastClose) + endRun + entriesXml.slice(lastClose);
}
