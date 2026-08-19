/**
 * Comments-extended descriptor — produces word/commentsExtended.xml
 * (Word 2013+ comment metadata: resolved state and reply threading).
 *
 * Each w15:commentEx links to a w:comment through the w14:paraId of the
 * comment's first paragraph, not through the w:comment/@w:id.
 *
 * Reference: ISO/IEC 29500-4 wml.xsd, CT_CommentsEx / CT_CommentEx
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, escapeXml } from "@office-open/xml";

import { COMMENTS_NS } from "./comments";

/** Options for one extended-comment entry (w15:commentEx). */
export interface CommentExtendedOptions {
  /** w14:paraId of the comment's first paragraph (w15:paraId, required). */
  paraId: string;
  /** w14:paraId of the first paragraph of the comment this one replies to (w15:paraIdParent). */
  paraIdParent?: string;
  /** Resolved state; Word writes both 0 and 1 explicitly (w15:done). */
  done?: boolean;
  /**
   * Emit as w15:commentExNG — the next-gen comment variant element, same
   * attribute set as w15:commentEx. Round-trip only.
   */
  nextGen?: true;
}

export const commentsExtendedDesc: CustomDescriptor<CommentExtendedOptions[]> = {
  kind: "custom",

  stringify(opts) {
    const parts: string[] = [`<w15:commentsEx ${COMMENTS_NS}>`];
    for (const ex of opts) {
      const attrs = [`w15:paraId="${escapeXml(ex.paraId)}"`];
      if (ex.paraIdParent !== undefined)
        attrs.push(`w15:paraIdParent="${escapeXml(ex.paraIdParent)}"`);
      if (ex.done !== undefined) attrs.push(`w15:done="${ex.done ? 1 : 0}"`);
      parts.push(`<w15:${ex.nextGen ? "commentExNG" : "commentEx"} ${attrs.join(" ")}/>`);
    }
    parts.push("</w15:commentsEx>");
    return parts.join("");
  },

  parse(el) {
    const entries: CommentExtendedOptions[] = [];
    for (const child of el.elements ?? []) {
      if (child.name !== "w15:commentEx" && child.name !== "w15:commentExNG") continue;
      const paraId = attr(child, "w15:paraId");
      if (paraId === undefined) continue;
      const ex: Partial<CommentExtendedOptions> = { paraId };
      const paraIdParent = attr(child, "w15:paraIdParent");
      if (paraIdParent !== undefined) ex.paraIdParent = paraIdParent;
      const done = attr(child, "w15:done");
      if (done !== undefined) ex.done = parseOnOff(done) ?? false;
      if (child.name === "w15:commentExNG") ex.nextGen = true;
      entries.push(ex as CommentExtendedOptions);
    }
    return entries;
  },
};
