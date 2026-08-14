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
import { attr, attrNum } from "@office-open/xml";
import { documentNamespaceAttributes } from "@parts/document/document-attributes";
import { stringifyParagraphInline } from "@parts/inline";
import type { ParagraphOptions } from "@parts/paragraph/paragraph";

import { parseParagraph } from "../../body";
import type { BodyContext, DocxReadContext } from "../../context";

// ── Input ──

/** System footnote (separator / continuationSeparator). Round-tripped verbatim. */
export interface FootnoteSeparator {
  id: number;
  paragraphs: (ParagraphOptions | string)[];
}

export interface FootnotesData {
  notes: Map<number, (ParagraphOptions | string)[]>;
  /**
   * Separator footnote — id + content round-tripped from the source so it stays
   * consistent with settings.footnotePr, which references this id.
   */
  separator?: FootnoteSeparator;
  /** Continuation separator footnote — id + content round-tripped from the source. */
  continuationSeparator?: FootnoteSeparator;
}

// ── Constants ──

const NS = documentNamespaceAttributes([
  "m",
  "mc",
  "o",
  "r",
  "v",
  "w",
  "w10",
  "w14",
  "w15",
  "wne",
  "wp",
  "wp14",
  "wpc",
  "wpg",
  "wpi",
  "wps",
]);

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

/** Render a system footnote from round-tripped id + content, or the spec default. */
function footnoteSystemNote(
  type: "separator" | "continuationSeparator",
  sep: FootnoteSeparator | undefined,
  fallback: string,
  ctx: BodyContext,
): string {
  if (!sep) return fallback;
  const inner = sep.paragraphs.map((p) => stringifyParagraphInline(p, ctx)).join("");
  return `<w:footnote w:type="${type}" w:id="${sep.id}">${inner}</w:footnote>`;
}

export const footnotesDesc: CustomDescriptor<FootnotesData, BodyContext> = {
  kind: "custom",

  stringify(data, ctx) {
    const parts: string[] = [`<w:footnotes ${NS} mc:Ignorable="w14 w15 wp14">`];

    // Separator / continuation separator: round-trip id + content verbatim so they
    // stay consistent with settings.footnotePr (which references these ids). Fall
    // back to spec defaults only for freshly generated documents.
    parts.push(footnoteSystemNote("separator", data.separator, SEPARATOR_FOOTNOTE, ctx));
    parts.push(
      footnoteSystemNote(
        "continuationSeparator",
        data.continuationSeparator,
        CONTINUATION_SEPARATOR_FOOTNOTE,
        ctx,
      ),
    );

    for (const [id, paragraphs] of data.notes) {
      parts.push(`<w:footnote w:id="${id}">`);
      for (const [i, para] of paragraphs.entries()) {
        const pXml = stringifyParagraphInline(para, ctx);
        if (i === 0) {
          // footnoteRef goes after the paragraph open tag, but after <w:pPr>
          // when present — CT_P requires pPr to be the first child.
          const pPrEnd = pXml.indexOf("</w:pPr>");
          if (pPrEnd !== -1) {
            const at = pPrEnd + "</w:pPr>".length;
            parts.push(pXml.slice(0, at) + FOOTNOTE_REF_RUN + pXml.slice(at));
          } else {
            const openIdx = pXml.indexOf(">");
            if (openIdx !== -1) {
              parts.push(pXml.slice(0, openIdx + 1) + FOOTNOTE_REF_RUN + pXml.slice(openIdx + 1));
            } else {
              parts.push(pXml);
            }
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

  parse(el, ctx) {
    const notes = new Map<number, (ParagraphOptions | string)[]>();
    let separator: FootnoteSeparator | undefined;
    let continuationSeparator: FootnoteSeparator | undefined;
    for (const child of el.elements ?? []) {
      if (child.name !== "w:footnote") continue;
      const id = attrNum(child, "w:id");
      if (id === undefined) continue;
      const type = attr(child, "w:type");

      const paragraphs: (ParagraphOptions | string)[] = [];
      for (const sub of child.elements ?? []) {
        if (sub.name === "w:p") {
          paragraphs.push(parseParagraph(sub, ctx as DocxReadContext));
        }
      }

      // System footnotes carry type="separator"/"continuationSeparator" — capture
      // their id + content so stringify round-trips them verbatim (settings
      // references these ids). Normal footnotes (type absent/"normal") go to notes.
      if (type === "separator") {
        separator = { id, paragraphs };
      } else if (type === "continuationSeparator") {
        continuationSeparator = { id, paragraphs };
      } else {
        notes.set(id, paragraphs);
      }
    }
    return { notes, separator, continuationSeparator } as FootnotesData;
  },
};
