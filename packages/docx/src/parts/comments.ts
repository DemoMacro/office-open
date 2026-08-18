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
import type { BookmarkStartOptions } from "@parts/paragraph/links/bookmark";
import type { ParagraphOptions } from "@parts/paragraph/paragraph";
import type { CommentOptions } from "@parts/paragraph/run/comment-run";
import { tableDesc } from "@parts/table/descriptor";
import type { TableOptions } from "@parts/table/table";

import { parseParagraph } from "../body";
import type { BodyContext } from "../context";
import type { DocxReadContext } from "../context";
import { documentNamespaceAttributes } from "./document/document-attributes";
import { buildBookmarkStartAttrs, buildMarkupRangeAttrs, stringifyParagraphInline } from "./inline";

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
    if (child !== null && typeof child === "object" && "table" in child) {
      parts.push(tableDesc.stringify(child.table, ctx) ?? "");
    } else if (child !== null && typeof child === "object" && "bookmarkStart" in child) {
      parts.push(`<w:bookmarkStart ${buildBookmarkStartAttrs(child.bookmarkStart)}/>`);
    } else if (child !== null && typeof child === "object" && "bookmarkEnd" in child) {
      parts.push(`<w:bookmarkEnd ${buildMarkupRangeAttrs(child.bookmarkEnd)}/>`);
    } else {
      parts.push(stringifyParagraphInline(child as string | ParagraphOptions, ctx));
    }
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

      const children: CommentOptions["children"] = [];
      for (const sub of child.elements ?? []) {
        if (sub.name === "w:p") {
          children.push(parseParagraph(sub, ctx as DocxReadContext));
        } else if (sub.name === "w:tbl") {
          children.push({ table: tableDesc.parse(sub, ctx as never) as TableOptions });
        } else if (sub.name === "w:bookmarkStart") {
          const id = attrNum(sub, "w:id");
          const name = attr(sub, "w:name");
          if (id !== undefined && name !== undefined) {
            const bs: BookmarkStartOptions = { id, name };
            const disp = attr(sub, "w:displacedByCustomXml");
            if (disp === "before" || disp === "after") bs.displacedByCustomXml = disp;
            const colFirst = attrNum(sub, "w:colFirst");
            if (colFirst !== undefined) bs.colFirst = colFirst;
            const colLast = attrNum(sub, "w:colLast");
            if (colLast !== undefined) bs.colLast = colLast;
            children.push({ bookmarkStart: bs });
          }
        } else if (sub.name === "w:bookmarkEnd") {
          const id = attrNum(sub, "w:id");
          if (id !== undefined) children.push({ bookmarkEnd: { id } });
        }
      }
      comment.children = children;
      comments.push(comment as CommentOptions);
    }
    return comments;
  },
};
