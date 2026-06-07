/**
 * Comment descriptors for PPTX.
 *
 * Handles both comment authors (p:cmAuthorLst) and
 * per-slide comment lists (p:cmLst).
 *
 * @module
 */

import type { AuthorEntry, CommentEntry } from "@file/comment";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { escapeXml } from "@office-open/xml";
import { attr, attrNum, findChild, textOf } from "@office-open/xml";

// ── Comment Authors ──

export const commentAuthorsDesc: CustomDescriptor<AuthorEntry[]> = {
  kind: "custom",

  stringify(authors, _ctx) {
    const parts: string[] = [
      '<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">',
    ];
    for (const a of authors) {
      parts.push(
        `<p:cmAuthor id="${a.id}" name="${escapeXml(a.name)}" initials="${escapeXml(a.initials)}" clrIdx="${a.clrIdx}" lastIdx="${a.lastIdx}"/>`,
      );
    }
    parts.push("</p:cmAuthorLst>");
    return parts.join("");
  },

  parse(el, _ctx) {
    const authors: AuthorEntry[] = [];
    for (const child of el.elements ?? []) {
      if (child.name !== "p:cmAuthor") continue;
      const id = attrNum(child, "id");
      const name = attr(child, "name");
      const initials = attr(child, "initials");
      const clrIdx = attrNum(child, "clrIdx");
      const lastIdx = attrNum(child, "lastIdx");
      if (id !== undefined && name && initials && clrIdx !== undefined && lastIdx !== undefined) {
        authors.push({ id, name, initials, clrIdx, lastIdx });
      }
    }
    return authors as AuthorEntry[];
  },
};

// ── Slide Comments ──

export const slideCommentsDesc: CustomDescriptor<CommentEntry[]> = {
  kind: "custom",

  stringify(comments, _ctx) {
    const parts: string[] = [
      '<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">',
    ];
    for (const c of comments) {
      const attrs: string[] = [`authorId="${c.authorId}"`, `idx="${c.idx}"`];
      if (c.date != null) attrs.push(`dt="${escapeXml(c.date)}"`);
      if (c.modified !== undefined) attrs.push(`mod="${c.modified ? 1 : 0}"`);
      parts.push(`<p:cm ${attrs.join(" ")}>`);
      parts.push(`<p:pos x="${c.x}" y="${c.y}"/>`);
      parts.push(`<p:text>${escapeXml(c.text)}</p:text>`);
      parts.push("</p:cm>");
    }
    parts.push("</p:cmLst>");
    return parts.join("");
  },

  parse(el, _ctx) {
    const comments: CommentEntry[] = [];
    for (const cm of el.elements ?? []) {
      if (cm.name !== "p:cm") continue;
      const authorId = attrNum(cm, "authorId") ?? 0;
      const idx = attrNum(cm, "idx") ?? 0;
      const date = attr(cm, "dt");
      const modifiedAttr = attr(cm, "mod");
      const pos = findChild(cm, "p:pos");
      const x = pos ? (attrNum(pos, "x") ?? 0) : 0;
      const y = pos ? (attrNum(pos, "y") ?? 0) : 0;
      const textEl = findChild(cm, "p:text");
      const text = textEl ? (textOf(textEl) ?? "") : "";
      const entry: CommentEntry = { authorId, idx, x, y, text };
      if (date !== undefined) entry.date = date;
      if (modifiedAttr !== undefined) entry.modified = modifiedAttr === "1";
      comments.push(entry);
    }
    return comments as CommentEntry[];
  },
};
