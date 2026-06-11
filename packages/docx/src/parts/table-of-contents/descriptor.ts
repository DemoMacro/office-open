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
): string {
  const instr = tocInstructionStr(options);
  const aliasAttr = alias ? ` w:val="${escapeXml(alias)}"` : "";

  // SDT properties: alias + docPartObj
  const sdtPr =
    `<w:sdtPr>` +
    `<w:alias${aliasAttr}/>` +
    `<w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj>` +
    `</w:sdtPr>`;

  // SDT content: begin paragraph (with field instruction) + end paragraph
  const content =
    `<w:sdtContent>` +
    `<w:p><w:r><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:cstheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cs="Times New Roman"/></w:rPr><w:fldChar w:fldCharType="begin" w:dirty="1"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>` +
    `<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>` +
    `</w:sdtContent>`;

  return `<w:sdt>${sdtPr}${content}</w:sdt>`;
}
