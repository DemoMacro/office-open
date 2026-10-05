/**
 * Extensible-comments descriptor — produces word/commentsExtensible.xml.
 *
 * Reference: Microsoft wml-cex-2018.xsd, CT_CommentsExtensible /
 * CT_CommentExtensible
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import type { LongHexNumber } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, escapeXml, stringify as stringifyXml } from "@office-open/xml";

import type { BodyContext } from "../context";
import { commentsNamespaceAttributes } from "./comments";

/** Options for one extensible-comment entry (w16cex:commentExtensible). */
export interface CommentExtensibleOptions {
  /** Durable comment id (w16cex:durableId, required). */
  durableId: LongHexNumber;
  /** UTC comment date (w16cex:dateUtc). */
  dateUtc?: string;
  /** Placeholder indicator used by modern Word comment experiences. */
  intelligentPlaceholder?: boolean;
  /** Raw w16:extLst inner XML for future Microsoft extensions. Round-trip only. */
  ext?: string;
}

export const commentsExtensibleDesc: CustomDescriptor<CommentExtensibleOptions[], BodyContext> = {
  kind: "custom",

  stringify(opts, ctx) {
    const entries = opts
      .map((entry) => {
        const attrs = [`w16cex:durableId="${escapeXml(entry.durableId)}"`];
        if (entry.dateUtc !== undefined) attrs.push(`w16cex:dateUtc="${escapeXml(entry.dateUtc)}"`);
        if (entry.intelligentPlaceholder !== undefined)
          attrs.push(`w16cex:intelligentPlaceholder="${entry.intelligentPlaceholder ? 1 : 0}"`);
        const ext = entry.ext === undefined ? "" : `<w16:extLst>${entry.ext}</w16:extLst>`;
        return `<w16cex:commentExtensible ${attrs.join(" ")}>${ext}</w16cex:commentExtensible>`;
      })
      .join("");
    return `<w16cex:commentsExtensible ${commentsNamespaceAttributes(ctx)}>${entries}</w16cex:commentsExtensible>`;
  },

  parse(el) {
    return (el.elements ?? []).flatMap((child) => {
      if (child.name !== "w16cex:commentExtensible") return [];
      const durableId = attr(child, "w16cex:durableId");
      if (durableId === undefined) return [];
      const entry: Partial<CommentExtensibleOptions> = { durableId };
      const dateUtc = attr(child, "w16cex:dateUtc");
      if (dateUtc !== undefined) entry.dateUtc = dateUtc;
      const placeholder = attr(child, "w16cex:intelligentPlaceholder");
      if (placeholder !== undefined)
        entry.intelligentPlaceholder = parseOnOff(placeholder) ?? false;
      const extLst = (child.elements ?? []).find((node) => node.name === "w16:extLst");
      if (extLst) entry.ext = stringifyXml(extLst);
      return [entry as CommentExtensibleOptions];
    });
  },
};
