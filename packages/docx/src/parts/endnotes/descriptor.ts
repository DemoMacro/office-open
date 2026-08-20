/**
 * Endnotes descriptor — produces word/endnotes.xml.
 *
 * Generates the complete `<w:endnotes>` element including:
 * - Separator endnote (id round-tripped from source; default -1)
 * - Continuation separator endnote (id round-tripped from source; default 0)
 * - User endnotes with auto-injected endnoteRef in first paragraph
 *
 * Reference: http://officeopenxml.com/WPfootnotes.php, CT_Endnotes / CT_FtnEdn
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import {
  createNotesDesc,
  type NoteChild,
  type NoteSeparator,
  type NotesData,
} from "@parts/notes/shared";

import type { BodyContext } from "../../context";

// ── Input (aliases of the shared notes types) ──

/** A user endnote. `id` is auto-assigned (1, 2, …) when omitted; round-tripped entries carry theirs. */
export interface EndnoteOptions {
  id?: number;
  children: NoteChild[];
}

/** System endnote (separator / continuationSeparator). Round-tripped verbatim. */
export type EndnoteSeparator = NoteSeparator;

export type EndnotesData = NotesData;

// ── Constants ──

/** XML for the endnoteRef run — auto-injected at start of first paragraph. */
const ENDNOTE_REF_RUN =
  '<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r>';

/** Default separator endnote (id=-1) for freshly generated documents. */
const SEPARATOR_ENDNOTE =
  '<w:endnote w:type="separator" w:id="-1">' +
  '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' +
  "<w:r><w:separator/></w:r>" +
  "</w:p>" +
  "</w:endnote>";

/** Default continuation separator endnote (id=0) for freshly generated documents. */
const CONTINUATION_SEPARATOR_ENDNOTE =
  '<w:endnote w:type="continuationSeparator" w:id="0">' +
  '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' +
  "<w:r><w:continuationSeparator/></w:r>" +
  "</w:p>" +
  "</w:endnote>";

// ── Descriptor ──

export const endnotesDesc: CustomDescriptor<EndnotesData, BodyContext> = createNotesDesc({
  rootTag: "w:endnotes",
  noteTag: "w:endnote",
  refRunXml: ENDNOTE_REF_RUN,
  separatorXml: SEPARATOR_ENDNOTE,
  continuationSeparatorXml: CONTINUATION_SEPARATOR_ENDNOTE,
  // endnoteRef goes right after the paragraph open tag (historical output).
  insertRefAfterParagraphProperties: false,
});
