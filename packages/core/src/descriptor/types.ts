/**
 * Descriptor type definitions for declarative XML mapping.
 *
 * A Descriptor describes the bidirectional mapping between a TypeScript
 * options object and its XML representation. The same descriptor drives
 * both stringify (Options → XML string) and parse (Element → Options).
 *
 * @module
 */

import type { Element as XmlElement } from "@office-open/xml";

import type { ReadContext, WriteContext } from "./context";

// ── Core descriptor types ──

/** Element descriptor — declarative XML mapping. */
export interface ElementDescriptor<T> {
  readonly kind: "element";
  readonly tag: string;
  readonly attrs?: readonly AttrSpec<T>[];
  readonly content?: readonly ContentSpec<T>[];
}

/** Custom descriptor — for complex logic that doesn't fit the declarative model. */
export interface CustomDescriptor<T> {
  readonly kind: "custom";
  stringify(value: T, ctx: WriteContext): string | undefined;
  parse(el: XmlElement, ctx: ReadContext): Partial<T>;
}

/** Union type for all descriptors. */
export type Descriptor<T> = ElementDescriptor<T> | CustomDescriptor<T>;

// ── Attribute specification ──

export interface AttrSpec<T> {
  /** Property key in the Options object. */
  readonly key: keyof T & string;
  /** XML attribute name (e.g. "w:val"). */
  readonly xmlName: string;
  /** Default value — omitted during stringify when equal. */
  readonly default?: unknown;
  /** Encode a JS value to an XML attribute string. Return undefined to skip. */
  readonly encode?: (v: any) => string | undefined;
  /** Decode an XML attribute string to a JS value. */
  readonly decode?: (raw: string) => any;
}

// ── Content specifications ──

/** Single child element mapped to a property. */
export interface ChildSpec<T> {
  readonly kind: "child";
  readonly key: keyof T & string;
  readonly tag: string;
  readonly desc: Descriptor<any>;
}

/** Multiple child elements of the same tag mapped to an array property. */
export interface ChildrenSpec<T> {
  readonly kind: "children";
  readonly key: keyof T & string;
  readonly tag: string;
  readonly desc: Descriptor<any>;
}

/** Union child — one of several possible child elements. */
export interface UnionSpec<T> {
  readonly kind: "union";
  readonly key: keyof T & string;
  readonly variants: readonly UnionVariant[];
}

/** Text content mapped to a property. */
export interface TextSpec<T> {
  readonly kind: "text";
  readonly key: keyof T & string;
}

export interface UnionVariant {
  readonly tag: string;
  readonly match: (opts: any) => boolean;
  readonly desc: Descriptor<any>;
}

/** Discriminated union of all content spec types. */
export type ContentSpec<T> =
  | ChildSpec<T>
  | ChildrenSpec<T>
  | UnionSpec<T>
  | TextSpec<T>
  | CustomDescriptor<T>;
