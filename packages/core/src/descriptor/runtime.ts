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
export function stringify<T>(
  desc: CustomDescriptor<T>,
  value: T,
  ctx: WriteContext,
): string | undefined {
  return desc.stringify(value, ctx);
}

// ── Read path ──

/** Parse an XML Element into an Options object using its descriptor. */
export function parse<T>(desc: CustomDescriptor<T>, el: XmlElement, ctx: ReadContext): T {
  return desc.parse(el, ctx);
}
