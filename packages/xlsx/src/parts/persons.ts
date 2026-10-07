/**
 * Threaded-comment persons part types and descriptor (xl/persons/person.xml).
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, children, escapeXml } from "@office-open/xml";

const TC_NS = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";
const PERSON_ATTRIBUTES = ["displayName", "userId", "providerId", "id"] as const;

/** Threaded-comment author (CT_Person). */
export interface PersonOptions {
  /** Author display name (required). */
  displayName: string;
  /** Author identity within the provider (optional). */
  userId?: string;
  /** Identity provider identifier (optional, default "None"). */
  providerId?: string;
  /** Person identifier (optional). */
  id?: string;
}

/** Options for xl/persons/person.xml (CT_Persons). */
export interface PersonsOptions {
  persons?: PersonOptions[];
}

/** Structured parse failure with the exact persons-part location. */
export class PersonsParseError extends Error {
  constructor(
    readonly part: string,
    readonly path: string,
    readonly name: string,
    readonly reason: string,
  ) {
    super(`${part}${path}: ${name}: ${reason}`);
  }
}

export const personsDesc: CustomDescriptor<PersonsOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const personsXml = (opts.persons ?? [])
      .map((person) => {
        const attrs = [
          ` displayName="${escapeXml(person.displayName)}"`,
          ...(person.userId !== undefined ? [` userId="${escapeXml(person.userId)}"`] : []),
          ...(person.providerId !== undefined
            ? [` providerId="${escapeXml(person.providerId)}"`]
            : []),
          ...(person.id !== undefined ? [` id="${escapeXml(person.id)}"`] : []),
        ].join("");
        return `<tc:person${attrs}/>`;
      })
      .join("");
    return `<tc:personList xmlns:tc="${TC_NS}">${personsXml}</tc:personList>`;
  },

  parse(el, _ctx) {
    const part = "xl/persons/person.xml";
    for (const child of el.elements ?? []) {
      if (child.name !== "tc:person") {
        throw new PersonsParseError(
          part,
          "/personList",
          child.name ?? "unknown",
          "unexpected child of personList",
        );
      }
    }
    const persons = children(el, "tc:person").map((element) => {
      const path = "/personList/person";
      const displayName = attr(element, "displayName");
      if (displayName === undefined) {
        throw new PersonsParseError(part, path, "@displayName", "missing required value");
      }
      for (const attribute of Object.keys(element.attributes ?? {})) {
        if (attribute === "xmlns:tc" || attribute === "xmlns") continue;
        if (!PERSON_ATTRIBUTES.includes(attribute as never)) {
          throw new PersonsParseError(part, path, `@${attribute}`, "unexpected attribute");
        }
      }
      return {
        displayName,
        ...(attr(element, "userId") !== undefined ? { userId: attr(element, "userId") } : {}),
        ...(attr(element, "providerId") !== undefined
          ? { providerId: attr(element, "providerId") }
          : {}),
        ...(attr(element, "id") !== undefined ? { id: attr(element, "id") } : {}),
      };
    });
    const result: PersonsOptions = {};
    result.persons = persons;
    return result;
  },
};
