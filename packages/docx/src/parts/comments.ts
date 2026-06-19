/**
 * Comments descriptor — produces word/comments.xml.
 *
 * Stringifies pure JSON CommentsOptions into XML without creating
 * Comment/Comments class instances.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrNum } from "@office-open/xml";
import type { ParagraphOptions } from "@parts/paragraph/paragraph";
import type { CommentsOptions, CommentOptions } from "@parts/paragraph/run/comment-run";

import { parseParagraph } from "../body";
import type { BodyContext } from "../context";
import type { DocxReadContext } from "../context";
import { stringifyParagraphInline } from "./inline";

// ── XML helpers ──

function escapeAttr(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

const COMMENTS_NS =
  'xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" ' +
  'xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" ' +
  'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" ' +
  'xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" ' +
  'xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" ' +
  'xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" ' +
  'xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" ' +
  'xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" ' +
  'xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" ' +
  'xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" ' +
  'xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" ' +
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w10="urn:schemas-microsoft-com:office:word" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
  'xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" ' +
  'xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" ' +
  'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" ' +
  'xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" ' +
  'xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" ' +
  'xmlns:wne="http://schemas.openxmlformats.org/office/word/2006/wordml" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" ' +
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
  'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';

// ── Comment stringification ──

function stringifyComment(opts: CommentOptions, ctx: BodyContext): string {
  const dateStr =
    typeof opts.date === "string" ? opts.date : (opts.date ?? new Date()).toISOString();
  // w:author is XSD-required (CT_TrackChange); default to empty string when absent.
  const attrs: string[] = [
    `w:id="${opts.id}"`,
    `w:author="${escapeAttr(opts.author ?? "")}"`,
    `w:date="${escapeAttr(dateStr)}"`,
  ];
  if (opts.initials !== undefined) attrs.push(`w:initials="${escapeAttr(opts.initials)}"`);

  const parts: string[] = [];
  for (const child of opts.children) {
    parts.push(stringifyParagraphInline(child as string | ParagraphOptions, ctx));
  }

  return `<w:comment ${attrs.join(" ")}>${parts.join("")}</w:comment>`;
}

// ── Descriptor ──

export const commentsDesc: CustomDescriptor<CommentsOptions, BodyContext> = {
  kind: "custom",

  stringify(opts, ctx) {
    const parts: string[] = [`<w:comments ${COMMENTS_NS}>`];

    for (const child of opts.children) {
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

      const children: unknown[] = [];
      for (const sub of child.elements ?? []) {
        if (sub.name === "w:p") {
          children.push(parseParagraph(sub, ctx as DocxReadContext));
        }
      }
      comment.children = children;
      comments.push(comment as CommentOptions);
    }
    return { children: comments };
  },
};
