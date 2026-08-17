/**
 * Comments descriptor — produces word/comments.xml.
 *
 * Stringifies pure JSON comment options into XML without creating
 * Comment/Comments class instances.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrNum, escapeXml } from "@office-open/xml";
import type { ParagraphOptions } from "@parts/paragraph/paragraph";
import type { CommentOptions } from "@parts/paragraph/run/comment-run";

import { parseParagraph } from "../body";
import type { BodyContext } from "../context";
import type { DocxReadContext } from "../context";
import { documentNamespaceAttributes } from "./document/document-attributes";
import { stringifyParagraphInline } from "./inline";

/** Root namespace header shared by the comment-infrastructure parts
 *  (comments.xml, people.xml, commentsExtended.xml). Word writes the same
 *  declaration set plus mc:Ignorable on all three roots. */
export const COMMENTS_NS =
  documentNamespaceAttributes([
    "aink",
    "am3d",
    "cx",
    "cx1",
    "cx2",
    "cx3",
    "cx4",
    "cx5",
    "cx6",
    "cx7",
    "cx8",
    "m",
    "mc",
    "o",
    "r",
    "v",
    "w",
    "w10",
    "w14",
    "w15",
    "w16",
    "w16cex",
    "w16cid",
    "w16sdtdh",
    "w16se",
    "wne",
    "wp",
    "wp14",
    "wpg",
    "wpi",
    "wps",
  ]) + ' mc:Ignorable="w14 w15 wp14"';

// ── Comment stringification ──

function stringifyComment(opts: CommentOptions, ctx: BodyContext): string {
  const dateStr = opts.date ?? new Date().toISOString();
  // w:author is XSD-required (CT_TrackChange); default to empty string when absent.
  const attrs: string[] = [
    `w:id="${opts.id}"`,
    `w:author="${escapeXml(opts.author ?? "")}"`,
    `w:date="${escapeXml(dateStr)}"`,
  ];
  if (opts.initials !== undefined) attrs.push(`w:initials="${escapeXml(opts.initials)}"`);

  const parts: string[] = [];
  for (const child of opts.children) {
    parts.push(stringifyParagraphInline(child, ctx));
  }

  return `<w:comment ${attrs.join(" ")}>${parts.join("")}</w:comment>`;
}

// ── Descriptor ──

export const commentsDesc: CustomDescriptor<CommentOptions[], BodyContext> = {
  kind: "custom",

  stringify(opts, ctx) {
    const parts: string[] = [`<w:comments ${COMMENTS_NS}>`];

    for (const child of opts) {
      parts.push(stringifyComment(child, ctx));
    }

    parts.push("</w:comments>");
    return parts.join("");
  },

  parse(el, ctx) {
    const comments: CommentOptions[] = [];
    for (const child of el.elements ?? []) {
      if (child.name !== "w:comment") continue;
      const id = attrNum(child, "w:id");
      if (id === undefined) continue;
      const comment: Partial<CommentOptions> = { id };
      const date = attr(child, "w:date");
      if (date) comment.date = date;
      const author = attr(child, "w:author");
      if (author !== undefined) comment.author = author;
      const initials = attr(child, "w:initials");
      if (initials !== undefined) comment.initials = initials;

      const children: (string | ParagraphOptions)[] = [];
      for (const sub of child.elements ?? []) {
        if (sub.name === "w:p") {
          children.push(parseParagraph(sub, ctx as DocxReadContext));
        }
      }
      comment.children = children;
      comments.push(comment as CommentOptions);
    }
    return comments;
  },
};
