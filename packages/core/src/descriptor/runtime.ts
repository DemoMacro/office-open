/**
 * Descriptor runtime: stringify (write) and parse (parse path) functions.
 *
 * Write path: Options → stringify(desc, opts, ctx) → string
 * Parse path: Element → parse(desc, el, ctx) → Partial<Options>
 *
 * No intermediate representation — each path is a single step.
 *
 * @module
 */

import { escapeXml } from "@office-open/xml";
import { findChild, children, textOf } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type { ReadContext, WriteContext } from "./context";
import type {
  AttrSpec,
  ChildrenSpec,
  ChildSpec,
  CustomDescriptor,
  ElementDescriptor,
  Descriptor,
  UnionSpec,
  TextSpec,
  ContentSpec,
} from "./types";

// ── Write path ──

/**
 * Serialize an Options object to an XML string using its descriptor.
 * Returns `undefined` when an optional element should be omitted.
 */
export function stringify<T>(desc: Descriptor<T>, value: T, ctx: WriteContext): string | undefined {
  if (desc.kind === "custom") return desc.stringify(value, ctx);
  return stringifyElement(desc, value, ctx);
}

function stringifyElement<T>(
  desc: ElementDescriptor<T>,
  value: T,
  ctx: WriteContext,
): string | undefined {
  const tag = desc.tag;

  // Build attribute string
  const attrStr = stringifyAttrs(desc.attrs, value);

  // Build child content
  let hasContent = false;
  const parts: string[] = [];

  if (desc.content) {
    for (let i = 0; i < desc.content.length; i++) {
      const spec = desc.content[i];
      const s = stringifyContentSpec(spec, value, ctx);
      if (s !== undefined) {
        hasContent = true;
        parts.push(s);
      }
    }
  }

  // Self-closing element with no meaningful content
  if (!hasContent) {
    if (!attrStr) return undefined;
    return `<${tag}${attrStr}/>`;
  }

  parts.unshift(`<${tag}${attrStr}>`);
  parts.push(`</${tag}>`);
  return parts.join("");
}

function stringifyAttrs<T>(attrs: readonly AttrSpec<T>[] | undefined, value: T): string {
  if (!attrs) return "";
  const parts: string[] = [];
  for (let i = 0; i < attrs.length; i++) {
    const spec = attrs[i];
    const raw = (value as Record<string, unknown>)[spec.key];
    if (raw === undefined) continue;

    // Skip default values
    if (spec.default !== undefined && raw === spec.default) continue;

    const encoded = spec.encode
      ? spec.encode(raw)
      : typeof raw === "string" || typeof raw === "number" || typeof raw === "boolean"
        ? String(raw)
        : String(raw as string | number);
    if (encoded === undefined) continue;

    parts.push(`${spec.xmlName}="${escapeXml(encoded)}"`);
  }
  return parts.length ? " " + parts.join(" ") : "";
}

function stringifyContentSpec<T>(
  spec: ContentSpec<T>,
  value: T,
  ctx: WriteContext,
): string | undefined {
  switch (spec.kind) {
    case "child":
      return stringifyChild(spec, value, ctx);
    case "children":
      return stringifyChildren(spec, value, ctx);
    case "union":
      return stringifyUnion(spec, value, ctx);
    case "text":
      return stringifyText(spec, value);
    case "custom":
      return stringifyCustom(spec, value, ctx);
  }
}

function stringifyChild<T>(spec: ChildSpec<T>, value: T, ctx: WriteContext): string | undefined {
  const childValue = (value as Record<string, unknown>)[spec.key];
  if (childValue === undefined || childValue === null) return undefined;
  return stringify(spec.desc, childValue, ctx);
}

function stringifyChildren<T>(
  spec: ChildrenSpec<T>,
  value: T,
  ctx: WriteContext,
): string | undefined {
  const items = (value as Record<string, unknown>)[spec.key] as readonly any[] | undefined;
  if (!items || items.length === 0) return undefined;

  const parts: string[] = [];
  for (let i = 0; i < items.length; i++) {
    const s = stringify(spec.desc, items[i], ctx);
    if (s !== undefined) parts.push(s);
  }
  return parts.length ? parts.join("") : undefined;
}

function stringifyUnion<T>(spec: UnionSpec<T>, value: T, ctx: WriteContext): string | undefined {
  const childValue = (value as Record<string, unknown>)[spec.key];
  if (childValue === undefined || childValue === null) return undefined;

  const variants = spec.variants;
  for (let i = 0; i < variants.length; i++) {
    const v = variants[i];
    if (v.match(childValue)) {
      return stringify(v.desc, childValue, ctx);
    }
  }
  return undefined;
}

function stringifyText<T>(spec: TextSpec<T>, value: T): string | undefined {
  const text = (value as Record<string, unknown>)[spec.key];
  if (text === undefined || text === null) return undefined;
  return escapeXml(typeof text === "string" ? text : String(text as string | number));
}

function stringifyCustom<T>(
  spec: CustomDescriptor<T>,
  value: T,
  ctx: WriteContext,
): string | undefined {
  return spec.stringify(value, ctx);
}

// ── Read path ──

/**
 * Parse an XML Element into an Options object using its descriptor.
 */
export function parse<T>(desc: Descriptor<T>, el: XmlElement, ctx: ReadContext): T {
  if (desc.kind === "custom") return desc.parse(el, ctx);
  return parseElement(desc, el, ctx);
}

function parseElement<T>(desc: ElementDescriptor<T>, el: XmlElement, ctx: ReadContext): T {
  const result: Record<string, unknown> = {};

  // Read attributes
  if (desc.attrs && el.attributes) {
    for (let i = 0; i < desc.attrs.length; i++) {
      const spec = desc.attrs[i];
      const raw = el.attributes[spec.xmlName];
      if (raw !== undefined) {
        result[spec.key] = spec.decode ? spec.decode(String(raw)) : raw;
      }
    }
  }

  // Read child content
  if (desc.content) {
    for (let i = 0; i < desc.content.length; i++) {
      const spec = desc.content[i];
      parseContentSpec(spec, el, ctx, result);
    }
  }

  return result as T;
}

function parseContentSpec<T>(
  spec: ContentSpec<T>,
  el: XmlElement,
  ctx: ReadContext,
  result: Record<string, unknown>,
): void {
  switch (spec.kind) {
    case "child": {
      const child = findChild(el, spec.tag);
      if (child) result[spec.key] = parse(spec.desc, child, ctx);
      break;
    }
    case "children": {
      const items = children(el, spec.tag);
      if (items.length) result[spec.key] = items.map((c) => parse(spec.desc, c, ctx));
      break;
    }
    case "union": {
      for (let i = 0; i < spec.variants.length; i++) {
        const v = spec.variants[i];
        const child = findChild(el, v.tag);
        if (child) {
          result[spec.key] = parse(v.desc, child, ctx);
          break;
        }
      }
      break;
    }
    case "text": {
      const text = textOf(el);
      if (text) result[spec.key] = text;
      break;
    }
    case "custom": {
      Object.assign(result, spec.parse(el, ctx));
      break;
    }
  }
}
