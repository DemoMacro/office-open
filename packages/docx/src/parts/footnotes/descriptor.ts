/**
 * Footnotes descriptor — produces word/footnotes.xml.
 *
 * Generates the complete `<w:footnotes>` element including:
 * - Separator footnote (id=-1)
 * - Continuation separator footnote (id=0)
 * - User footnotes with auto-injected footnoteRef in first paragraph
 *
 * Reference: http://officeopenxml.com/WPfootnotes.php, CT_Footnotes / CT_FtnEdn
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { stringifyParagraphInline } from "@parts/inline";
import type { ParagraphOptions } from "@parts/paragraph/paragraph";

import type { BodyContext } from "../../context";

// ── Input ──

export interface FootnotesData {
  notes: Map<number, (ParagraphOptions | string)[]>;
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

/** XML for the footnoteRef run — auto-injected at start of first paragraph. */
const FOOTNOTE_REF_RUN =
  '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>';

/** Separator footnote (id=-1). */
const SEPARATOR_FOOTNOTE =
  '<w:footnote w:type="separator" w:id="-1">' +
  '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' +
  "<w:r><w:separator/></w:r>" +
  "</w:p>" +
  "</w:footnote>";

/** Continuation separator footnote (id=0). */
const CONTINUATION_SEPARATOR_FOOTNOTE =
  '<w:footnote w:type="continuationSeparator" w:id="0">' +
  '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' +
  "<w:r><w:continuationSeparator/></w:r>" +
  "</w:p>" +
  "</w:footnote>";

// ── Descriptor ──

export const footnotesDesc: CustomDescriptor<FootnotesData, BodyContext> = {
  kind: "custom",

  stringify(data, ctx) {
    const parts: string[] = [
      `<w:footnotes ${NS} mc:Ignorable="w14 w15 wp14">`,
      SEPARATOR_FOOTNOTE,
      CONTINUATION_SEPARATOR_FOOTNOTE,
    ];

    for (const [id, paragraphs] of data.notes) {
      parts.push(`<w:footnote w:id="${id}">`);
      for (let i = 0; i < paragraphs.length; i++) {
        const pXml = stringifyParagraphInline(paragraphs[i], ctx);
        if (i === 0) {
          // Insert footnoteRef after <w:p> or <w:p ...>
          const openIdx = pXml.indexOf(">");
          if (openIdx !== -1) {
            parts.push(pXml.slice(0, openIdx + 1) + FOOTNOTE_REF_RUN + pXml.slice(openIdx + 1));
          } else {
            parts.push(pXml);
          }
        } else {
          parts.push(pXml);
        }
      }
      parts.push("</w:footnote>");
    }

    parts.push("</w:footnotes>");
    return parts.join("");
  },

  parse(_el, _ctx) {
    return {};
  },
};
