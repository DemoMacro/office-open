/**
 * Descriptor type definitions.
 *
 * A Descriptor describes the bidirectional mapping between a TypeScript
 * options object and its XML representation. The same descriptor drives both
 * stringify (Options → XML string) and parse (Element → Options).
 *
 * Every descriptor is a CustomDescriptor whose stringify()/parse() methods are
 * hand-written for the part.
 *
 * @module
 */

import type { Element as XmlElement } from "@office-open/xml";

import type { ReadContext, WriteContext } from "./context";

/** Custom descriptor — hand-written stringify/parse for an OOXML part. */
export interface CustomDescriptor<T, Ctx = WriteContext> {
  kind: "custom";
  stringify(value: T, ctx: Ctx): string | undefined;
  parse(el: XmlElement, ctx: ReadContext): T;
}

/** Alias for "any descriptor" — retained for call-site readability. */
export type Descriptor<T> = CustomDescriptor<T>;
