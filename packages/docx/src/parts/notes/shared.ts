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
import type { Element } from "@office-open/xml";
import { documentNamespaceAttributes } from "@parts/document/document-attributes";
import type { ParagraphOptions } from "@parts/paragraph/paragraph";
import type { SectionChild } from "@shared/section";

import type { BodyContext, DocxReadContext } from "../../context";

export type NoteChild = SectionChild | ParagraphOptions | string;

/** System note (separator / continuationSeparator). Round-tripped verbatim. */
export interface NoteSeparator {
  id: number;
  children: SectionChild[];
}

export interface NotesData {
  notes: Map<number, NoteChild[]>;
  /**
   * Separator note — id + content round-tripped from the source so it stays
   * consistent with settings.footnoteProperties/endnoteProperties, which reference this id.
   * null = the parsed part carried no separator note (emit nothing; only fresh
   * data falls back to the spec default).
   */
  separator?: NoteSeparator | null;
  /** Continuation separator note — same states as separator. */
  continuationSeparator?: NoteSeparator | null;
  /** Continuation notice note (w:type="continuationNotice") — same states. */
  continuationNotice?: NoteSeparator | null;
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

type ParseNoteChild = (el: Element, ctx: DocxReadContext) => SectionChild;

let parseNoteChild: ParseNoteChild | undefined;

/** Register the block-level child parser without introducing a notes↔body cycle. */
export function setNotesParseChild(fn: ParseNoteChild): void {
  parseNoteChild = fn;
}

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
    type: "separator" | "continuationSeparator" | "continuationNotice",
    sep: NoteSeparator | null | undefined,
    fallback: string,
    ctx: BodyContext,
  ): string => {
    // null = the parsed source carried no such system note — emit nothing.
    if (sep === null) return "";
    if (!sep) return fallback;
    const inner = sep.children.map((child) => ctx.stringifyChild(child)).join("");
    return `<${cfg.noteTag} w:type="${type}" w:id="${sep.id}">${inner}</${cfg.noteTag}>`;
  };

  return {
    kind: "custom",

    stringify(data, ctx) {
      const parts: string[] = [`<${cfg.rootTag} ${NS} mc:Ignorable="w14 w15 wp14">`];

      // Separator / continuation separator: round-trip id + content verbatim so
      // they stay consistent with settings (which references these ids). Fall
      // back to spec defaults only for freshly generated documents. The
      // continuation notice has no spec default — it only ever round-trips.
      parts.push(systemNote("separator", data.separator, cfg.separatorXml, ctx));
      parts.push(
        systemNote(
          "continuationSeparator",
          data.continuationSeparator,
          cfg.continuationSeparatorXml,
          ctx,
        ),
      );
      if (data.continuationNotice) {
        parts.push(systemNote("continuationNotice", data.continuationNotice, "", ctx));
      }

      for (const [id, children] of data.notes) {
        parts.push(`<${cfg.noteTag} w:id="${id}">`);
        for (const [i, child] of children.entries()) {
          const isSectionChild =
            typeof child === "object" &&
            child !== null &&
            ("paragraph" in child ||
              "table" in child ||
              "toc" in child ||
              "textbox" in child ||
              "sdt" in child ||
              "altChunk" in child ||
              "customXml" in child ||
              "bookmarkStart" in child ||
              "bookmarkEnd" in child ||
              "rawXml" in child);
          const isParagraphInput = !isSectionChild || (isSectionChild && "paragraph" in child);
          const sectionChild: SectionChild = isSectionChild
            ? child
            : { paragraph: child as ParagraphOptions | string };
          const childXml = ctx.stringifyChild(sectionChild);
          // The conventional reference mark belongs only in a first paragraph.
          // A note beginning with a table or another block is valid and must not
          // receive an invalid run outside w:p.
          const hasRefMark =
            childXml.includes("<w:footnoteRef/>") || childXml.includes("<w:endnoteRef/>");
          if (i === 0 && isParagraphInput && !hasRefMark && !isEmptyParagraph(childXml)) {
            parts.push(injectRefRun(childXml, cfg));
          } else {
            parts.push(childXml);
          }
        }
        parts.push(`</${cfg.noteTag}>`);
      }

      parts.push(`</${cfg.rootTag}>`);
      return parts.join("");
    },

    parse(el, ctx) {
      const notes = new Map<number, NoteChild[]>();
      let separator: NoteSeparator | undefined;
      let continuationSeparator: NoteSeparator | undefined;
      let continuationNotice: NoteSeparator | undefined;
      for (const child of el.elements ?? []) {
        if (child.name !== cfg.noteTag) continue;
        const id = attrNum(child, "w:id");
        if (id === undefined) continue;
        const type = attr(child, "w:type");

        const children: SectionChild[] = [];
        for (const sub of child.elements ?? []) {
          if (sub.type !== "element") continue;
          if (!parseNoteChild) throw new Error("Notes child parser is not registered");
          children.push(parseNoteChild(sub, ctx as DocxReadContext));
        }

        // System notes carry a w:type (separator/continuationSeparator/
        // continuationNotice — the ST_FtnEdn system members) — capture their
        // id + content so stringify round-trips them verbatim (settings
        // references these ids). Normal notes (type absent/"normal") go to notes.
        if (type === "separator") {
          separator = { id, children };
        } else if (type === "continuationSeparator") {
          continuationSeparator = { id, children };
        } else if (type === "continuationNotice") {
          continuationNotice = { id, children };
        } else {
          notes.set(id, children);
        }
      }
      return { notes, separator, continuationSeparator, continuationNotice } as NotesData;
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

/** A paragraph element with no children at all — self-closed or open+close. */
function isEmptyParagraph(pXml: string): boolean {
  return /^<[^/>]+\/>$/.test(pXml) || /^<[^>]+><\/[^>]+>$/.test(pXml);
}
