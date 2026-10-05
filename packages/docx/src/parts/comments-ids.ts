/**
 * Comments-IDs descriptor — produces word/commentsIds.xml.
 *
 * Word uses durable ids to link modern comment UI state across sessions. The
 * association with comments.xml remains through each comment's w14:paraId.
 *
 * Reference: Microsoft wml-cid-2016.xsd, CT_CommentsIds / CT_CommentId
 *
 * @module
 */

import type { LongHexNumber } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, escapeXml } from "@office-open/xml";

import type { BodyContext } from "../context";
import { commentsNamespaceAttributes } from "./comments";

/** Options for one durable comment identifier (w16cid:commentId). */
export interface CommentIdOptions {
  /** w14:paraId of the comment's first paragraph (w16cid:paraId, required). */
  paraId: LongHexNumber;
  /** Word's durable identifier for the comment (w16cid:durableId, required). */
  durableId: LongHexNumber;
}

export const commentsIdsDesc: CustomDescriptor<CommentIdOptions[], BodyContext> = {
  kind: "custom",

  stringify(opts, ctx) {
    const entries = opts
      .map(
        (entry) =>
          `<w16cid:commentId w16cid:paraId="${escapeXml(entry.paraId)}" w16cid:durableId="${escapeXml(entry.durableId)}"/>`,
      )
      .join("");
    return `<w16cid:commentsIds ${commentsNamespaceAttributes(ctx)}>${entries}</w16cid:commentsIds>`;
  },

  parse(el) {
    return (el.elements ?? []).flatMap((child) => {
      if (child.name !== "w16cid:commentId") return [];
      const paraId = attr(child, "w16cid:paraId");
      const durableId = attr(child, "w16cid:durableId");
      return paraId !== undefined && durableId !== undefined ? [{ paraId, durableId }] : [];
    });
  },
};
