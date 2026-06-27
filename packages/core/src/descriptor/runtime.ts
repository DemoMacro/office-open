/**
 * Descriptor runtime: stringify (write) and parse (parse path) functions.
 *
 * Write path: Options → stringify(desc, opts, ctx) → string
 * Parse path: Element → parse(desc, el, ctx) → Options
 *
 * Both paths delegate to the descriptor's own stringify()/parse() methods —
 * every descriptor is a CustomDescriptor.
 *
 * @module
 */

import type { Element as XmlElement } from "@office-open/xml";

import type { ReadContext, WriteContext } from "./context";
import type { CustomDescriptor } from "./types";

// ── Write path ──

/**
 * Serialize an Options object to an XML string using its descriptor.
 * Returns `undefined` when an optional element should be omitted.
 */
export function stringify<TInput, Ctx = WriteContext, TOutput = TInput>(
  desc: CustomDescriptor<TInput, Ctx, TOutput>,
  value: TInput,
  ctx: Ctx,
): string | undefined {
  return desc.stringify(value, ctx);
}

// ── Read path ──

/** Parse an XML Element into an Options object using its descriptor. */
export function parse<TInput, TOutput = TInput, Ctx = WriteContext>(
  desc: CustomDescriptor<TInput, Ctx, TOutput>,
  el: XmlElement,
  ctx: ReadContext,
): TOutput {
  return desc.parse(el, ctx);
}
