/**
 * Fluent builder for constructing ElementDescriptor instances.
 *
 * Usage:
 * ```ts
 * const myDesc = element<MyOptions>("w:p")
 *   .attr("alignment", "w:val")
 *   .child("spacing", "w:spacing", spacingDesc)
 *   .children("runs", "w:r", runDesc)
 *   .text("text")
 *   .build();
 * ```
 *
 * @module
 */

import type {
  AttrSpec,
  ChildrenSpec,
  ChildSpec,
  ContentSpec,
  CustomDescriptor,
  ElementDescriptor,
  UnionSpec,
  UnionVariant,
  TextSpec,
} from "./types";

export function element<T extends object>(tag: string): DescriptorBuilder<T> {
  return new DescriptorBuilder<T>(tag);
}

export class DescriptorBuilder<T extends object> {
  private readonly _tag: string;
  private readonly _attrs: AttrSpec<T>[] = [];
  private readonly _content: ContentSpec<T>[] = [];

  public constructor(tag: string) {
    this._tag = tag;
  }

  /** Add an attribute mapping. */
  public attr(
    key: keyof T & string,
    xmlName: string,
    opts?: {
      readonly default?: unknown;
      readonly encode?: (v: any) => string | undefined;
      readonly decode?: (raw: string) => any;
    },
  ): this {
    this._attrs.push({
      kind: "child" as never, // not needed for AttrSpec
      key,
      xmlName,
      ...opts,
    } as AttrSpec<T>);
    return this;
  }

  /** Add a single child element mapping. */
  public child(key: keyof T & string, tag: string, desc: ChildSpec<T>["desc"]): this {
    this._content.push({ kind: "child", key, tag, desc } as ChildSpec<T>);
    return this;
  }

  /** Add a repeating child element mapping. */
  public children(key: keyof T & string, tag: string, desc: ChildrenSpec<T>["desc"]): this {
    this._content.push({ kind: "children", key, tag, desc } as ChildrenSpec<T>);
    return this;
  }

  /** Add a union (one-of-several) child mapping. */
  public union(key: keyof T & string, variants: readonly UnionVariant[]): this {
    this._content.push({ kind: "union", key, variants } as UnionSpec<T>);
    return this;
  }

  /** Add a text content mapping. */
  public text(key: keyof T & string): this {
    this._content.push({ kind: "text", key } as TextSpec<T>);
    return this;
  }

  /** Add a custom content handler. */
  public custom(spec: CustomDescriptor<T>): this {
    this._content.push(spec);
    return this;
  }

  /** Build the immutable ElementDescriptor. */
  public build(): ElementDescriptor<T> {
    const result: ElementDescriptor<T> = {
      kind: "element",
      tag: this._tag,
    };
    if (this._attrs.length) (result as any).attrs = Object.freeze(this._attrs);
    if (this._content.length) (result as any).content = Object.freeze(this._content);
    return Object.freeze(result);
  }
}
