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
import { attr, attrNum } from "@office-open/xml";
import { stringifyParagraphInline } from "@parts/inline";
import type { ParagraphOptions } from "@parts/paragraph/paragraph";

import { parseParagraph } from "../../body";
import type { BodyContext, DocxReadContext } from "../../context";

// ── Input ──

/** System endnote (separator / continuationSeparator). Round-tripped verbatim. */
export interface EndnoteSeparator {
  id: number;
  paragraphs: (ParagraphOptions | string)[];
}

export interface EndnotesData {
  notes: Map<number, (ParagraphOptions | string)[]>;
  /**
   * Separator endnote — id + content round-tripped from the source so it stays
   * consistent with settings.endnotePr, which references this id.
   */
  separator?: EndnoteSeparator;
  /** Continuation separator endnote — id + content round-tripped from the source. */
  continuationSeparator?: EndnoteSeparator;
}

// ── Constants ──

const NS =
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w10="urn:schemas-microsoft-com:office:word" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
  'xmlns:wne="http://schemas.openxmlformats.org/office/word/2006/wordml" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" ' +
  'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" ' +
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
  'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';

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

/** Render a system endnote from round-tripped id + content, or the spec default. */
function endnoteSystemNote(
  type: "separator" | "continuationSeparator",
  sep: EndnoteSeparator | undefined,
  fallback: string,
  ctx: BodyContext,
): string {
  if (!sep) return fallback;
  const inner = sep.paragraphs.map((p) => stringifyParagraphInline(p, ctx)).join("");
  return `<w:endnote w:type="${type}" w:id="${sep.id}">${inner}</w:endnote>`;
}

export const endnotesDesc: CustomDescriptor<EndnotesData, BodyContext> = {
  kind: "custom",

  stringify(data, ctx) {
    const parts: string[] = [`<w:endnotes ${NS} mc:Ignorable="w14 w15 wp14">`];

    // Separator / continuation separator: round-trip id + content verbatim so they
    // stay consistent with settings.endnotePr (which references these ids). Fall
    // back to spec defaults only for freshly generated documents.
    parts.push(endnoteSystemNote("separator", data.separator, SEPARATOR_ENDNOTE, ctx));
    parts.push(
      endnoteSystemNote(
        "continuationSeparator",
        data.continuationSeparator,
        CONTINUATION_SEPARATOR_ENDNOTE,
        ctx,
      ),
    );

    for (const [id, paragraphs] of data.notes) {
      parts.push(`<w:endnote w:id="${id}">`);
      for (let i = 0; i < paragraphs.length; i++) {
        const pXml = stringifyParagraphInline(paragraphs[i], ctx);
        if (i === 0) {
          // Insert endnoteRef after <w:p> or <w:p ...>
          const openIdx = pXml.indexOf(">");
          if (openIdx !== -1) {
            parts.push(pXml.slice(0, openIdx + 1) + ENDNOTE_REF_RUN + pXml.slice(openIdx + 1));
          } else {
            parts.push(pXml);
          }
        } else {
          parts.push(pXml);
        }
      }
      parts.push("</w:endnote>");
    }

    parts.push("</w:endnotes>");
    return parts.join("");
  },

  parse(el, ctx) {
    const notes = new Map<number, (ParagraphOptions | string)[]>();
    let separator: EndnoteSeparator | undefined;
    let continuationSeparator: EndnoteSeparator | undefined;
    for (const child of el.elements ?? []) {
      if (child.name !== "w:endnote") continue;
      const id = attrNum(child, "w:id");
      if (id === undefined) continue;
      const type = attr(child, "w:type");

      const paragraphs: (ParagraphOptions | string)[] = [];
      for (const sub of child.elements ?? []) {
        if (sub.name === "w:p") {
          paragraphs.push(parseParagraph(sub, ctx as DocxReadContext));
        }
      }

      // System endnotes carry type="separator"/"continuationSeparator" — capture
      // their id + content so stringify round-trips them verbatim (settings
      // references these ids). Normal endnotes (type absent/"normal") go to notes.
      if (type === "separator") {
        separator = { id, paragraphs };
      } else if (type === "continuationSeparator") {
        continuationSeparator = { id, paragraphs };
      } else {
        notes.set(id, paragraphs);
      }
    }
    return { notes, separator, continuationSeparator } as EndnotesData;
  },
};
