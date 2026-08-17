/**
 * People descriptor — produces word/people.xml (Word 2013+ comment authors).
 *
 * Each w15:person carries the display author plus an optional contact
 * address; Word pairs them with comments by author string equality.
 *
 * Reference: ISO/IEC 29500-4 wml.xsd, CT_People / CT_Person
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, escapeXml } from "@office-open/xml";

import { COMMENTS_NS } from "./comments";

/** Options for one person entry (w15:person) — a comment participant,
 *  distinct from the bibliography's literature-author PersonOptions. */
export interface CommentPersonOptions {
  /** Author display name exactly as it appears on w:comment/@w:author (w15:author, required). */
  author: string;
  /** Contact address, usually an email (w15:contact). */
  contact?: string;
}

export const peopleDesc: CustomDescriptor<CommentPersonOptions[]> = {
  kind: "custom",

  stringify(opts) {
    const parts: string[] = [`<w15:people ${COMMENTS_NS}>`];
    for (const person of opts) {
      const attrs = [`w15:author="${escapeXml(person.author)}"`];
      if (person.contact !== undefined) attrs.push(`w15:contact="${escapeXml(person.contact)}"`);
      parts.push(`<w15:person ${attrs.join(" ")}/>`);
    }
    parts.push("</w15:people>");
    return parts.join("");
  },

  parse(el) {
    const people: CommentPersonOptions[] = [];
    for (const child of el.elements ?? []) {
      if (child.name !== "w15:person") continue;
      const author = attr(child, "w15:author");
      if (author === undefined) continue;
      const person: Partial<CommentPersonOptions> = { author };
      const contact = attr(child, "w15:contact");
      if (contact !== undefined) person.contact = contact;
      people.push(person as CommentPersonOptions);
    }
    return people;
  },
};
