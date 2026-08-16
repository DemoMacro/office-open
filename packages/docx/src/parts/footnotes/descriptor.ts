/**
 * Footnotes descriptor — produces word/footnotes.xml.
 *
 * Generates the complete `<w:footnotes>` element including:
 * - Separator footnote (id round-tripped from source; default -1)
 * - Continuation separator footnote (id round-tripped from source; default 0)
 * - User footnotes with auto-injected footnoteRef in first paragraph
 *
 * Reference: http://officeopenxml.com/WPfootnotes.php, CT_Footnotes / CT_FtnEdn
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { createNotesDesc, type NoteSeparator, type NotesData } from "@parts/notes/shared";
import type { ParagraphOptions } from "@parts/paragraph/paragraph";

import type { BodyContext } from "../../context";

// ── Input (aliases of the shared notes types) ──

/** A user footnote. `id` is auto-assigned (1, 2, …) when omitted; round-tripped entries carry theirs. */
export interface FootnoteOptions {
  id?: number;
  children: (ParagraphOptions | string)[];
}

/** System footnote (separator / continuationSeparator). Round-tripped verbatim. */
export type FootnoteSeparator = NoteSeparator;

export type FootnotesData = NotesData;

// ── Constants ──

/** XML for the footnoteRef run — auto-injected at start of first paragraph. */
const FOOTNOTE_REF_RUN =
  '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>';

/** Default separator footnote (id=-1) for freshly generated documents. */
const SEPARATOR_FOOTNOTE =
  '<w:footnote w:type="separator" w:id="-1">' +
  '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' +
  "<w:r><w:separator/></w:r>" +
  "</w:p>" +
  "</w:footnote>";

/** Default continuation separator footnote (id=0) for freshly generated documents. */
const CONTINUATION_SEPARATOR_FOOTNOTE =
  '<w:footnote w:type="continuationSeparator" w:id="0">' +
  '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' +
  "<w:r><w:continuationSeparator/></w:r>" +
  "</w:p>" +
  "</w:footnote>";

// ── Descriptor ──

export const footnotesDesc: CustomDescriptor<FootnotesData, BodyContext> = createNotesDesc({
  rootTag: "w:footnotes",
  noteTag: "w:footnote",
  refRunXml: FOOTNOTE_REF_RUN,
  separatorXml: SEPARATOR_FOOTNOTE,
  continuationSeparatorXml: CONTINUATION_SEPARATOR_FOOTNOTE,
  // footnoteRef goes after the paragraph open tag, but after <w:pPr> when
  // present — CT_P requires pPr to be the first child.
  insertRefAfterParagraphProperties: true,
});
