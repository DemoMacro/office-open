/**
 * Shared skeleton for the footnotes/endnotes parts.
 *
 * Both parts serialize the same CT_FtnEdn structure — only the element names,
 * the reference-run flavor (footnoteRef/endnoteRef), and the reference-run
 * insertion point differ: a footnote reference goes after <w:pPr> when present
 * (CT_P ordering), an endnote reference goes right after the paragraph open
 * tag, preserving each format's historical output byte-for-byte.
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

/** System note (separator / continuationSeparator). Round-tripped verbatim. */
export interface NoteSeparator {
  id: number;
  paragraphs: (ParagraphOptions | string)[];
}

export interface NotesData {
  notes: Map<number, (ParagraphOptions | string)[]>;
  /**
   * Separator note — id + content round-tripped from the source so it stays
   * consistent with settings.footnoteProperties/endnoteProperties, which reference this id.
   */
  separator?: NoteSeparator;
  /** Continuation separator note — id + content round-tripped from the source. */
  continuationSeparator?: NoteSeparator;
}

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

export interface NotesDescConfig {
  /** Root element name (`w:footnotes` / `w:endnotes`). */
  rootTag: string;
  /** Note element name (`w:footnote` / `w:endnote`). */
  noteTag: string;
  /** Reference run injected at the start of each note's first paragraph. */
  refRunXml: string;
  /** Separator note emitted when no round-tripped separator is present. */
  separatorXml: string;
  /** Continuation separator note emitted when no round-tripped one is present. */
  continuationSeparatorXml: string;
  /** Inject the reference run after <w:pPr> when present (CT_P ordering). */
  insertRefAfterParagraphProperties: boolean;
}

/** Build the complete footnotes/endnotes descriptor from a format config. */
export function createNotesDesc(cfg: NotesDescConfig): CustomDescriptor<NotesData, BodyContext> {
  /** Render a system note from round-tripped id + content, or the spec default. */
  const systemNote = (
    type: "separator" | "continuationSeparator",
    sep: NoteSeparator | undefined,
    fallback: string,
    ctx: BodyContext,
  ): string => {
    if (!sep) return fallback;
    const inner = sep.paragraphs.map((p) => stringifyParagraphInline(p, ctx)).join("");
    return `<${cfg.noteTag} w:type="${type}" w:id="${sep.id}">${inner}</${cfg.noteTag}>`;
  };

  return {
    kind: "custom",

    stringify(data, ctx) {
      const parts: string[] = [`<${cfg.rootTag} ${NS} mc:Ignorable="w14 w15 wp14">`];

      // Separator / continuation separator: round-trip id + content verbatim so
      // they stay consistent with settings (which references these ids). Fall
      // back to spec defaults only for freshly generated documents.
      parts.push(systemNote("separator", data.separator, cfg.separatorXml, ctx));
      parts.push(
        systemNote(
          "continuationSeparator",
          data.continuationSeparator,
          cfg.continuationSeparatorXml,
          ctx,
        ),
      );

      for (const [id, paragraphs] of data.notes) {
        parts.push(`<${cfg.noteTag} w:id="${id}">`);
        for (const [i, para] of paragraphs.entries()) {
          const pXml = stringifyParagraphInline(para, ctx);
          // Inject the reference run only on fresh content — a round-tripped
          // note keeps the parsed ref-mark run (with its own rPr) in place, and
          // injecting the template too would emit the mark twice.
          const hasRefMark = pXml.includes("<w:footnoteRef/>") || pXml.includes("<w:endnoteRef/>");
          if (i === 0 && !hasRefMark) {
            parts.push(injectRefRun(pXml, cfg));
          } else {
            parts.push(pXml);
          }
        }
        parts.push(`</${cfg.noteTag}>`);
      }

      parts.push(`</${cfg.rootTag}>`);
      return parts.join("");
    },

    parse(el, ctx) {
      const notes = new Map<number, (ParagraphOptions | string)[]>();
      let separator: NoteSeparator | undefined;
      let continuationSeparator: NoteSeparator | undefined;
      for (const child of el.elements ?? []) {
        if (child.name !== cfg.noteTag) continue;
        const id = attrNum(child, "w:id");
        if (id === undefined) continue;
        const type = attr(child, "w:type");

        const paragraphs: (ParagraphOptions | string)[] = [];
        for (const sub of child.elements ?? []) {
          if (sub.name === "w:p") {
            paragraphs.push(parseParagraph(sub, ctx as DocxReadContext));
          }
        }

        // System notes carry type="separator"/"continuationSeparator" — capture
        // their id + content so stringify round-trips them verbatim (settings
        // references these ids). Normal notes (type absent/"normal") go to notes.
        if (type === "separator") {
          separator = { id, paragraphs };
        } else if (type === "continuationSeparator") {
          continuationSeparator = { id, paragraphs };
        } else {
          notes.set(id, paragraphs);
        }
      }
      return { notes, separator, continuationSeparator } as NotesData;
    },
  };
}

/** Insert the reference run into the first paragraph's XML. */
function injectRefRun(pXml: string, cfg: NotesDescConfig): string {
  if (cfg.insertRefAfterParagraphProperties) {
    const pPrEnd = pXml.indexOf("</w:pPr>");
    if (pPrEnd !== -1) {
      const at = pPrEnd + "</w:pPr>".length;
      return pXml.slice(0, at) + cfg.refRunXml + pXml.slice(at);
    }
  }
  const openIdx = pXml.indexOf(">");
  return openIdx !== -1
    ? pXml.slice(0, openIdx + 1) + cfg.refRunXml + pXml.slice(openIdx + 1)
    : pXml;
}
