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

/**
 * Custom descriptor — hand-written stringify/parse for an OOXML part.
 *
 * `TInput` is the stringify input (the options shape callers build); `TOutput`
 * is the parse output (the shape parse produces). They default to the same
 * type, so the common case — a descriptor whose stringify and parse share one
 * options shape — is written as `CustomDescriptor<T>` unchanged. Only
 * descriptors whose parse yields a genuinely different shape (e.g. an
 * accumulator-based stringify vs. a structured parse result) pass a third
 * argument, which lets parse return its real type instead of lying via `as`.
 *
 * `Ctx` stays the second parameter to preserve the existing
 * `CustomDescriptor<T, BodyContext>` call sites that customize the write
 * context; `TOutput` is third so it never collides with that two-arg form.
 */
export interface CustomDescriptor<TInput, Ctx = WriteContext, TOutput = TInput> {
  kind: "custom";
  stringify(value: TInput, ctx: Ctx): string | undefined;
  parse(el: XmlElement, ctx: ReadContext): TOutput;
}

/** Alias for "any descriptor" — retained for call-site readability. */
export type Descriptor<TInput, Ctx = WriteContext, TOutput = TInput> = CustomDescriptor<
  TInput,
  Ctx,
  TOutput
>;
