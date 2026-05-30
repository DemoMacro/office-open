/**
 * XML attribute components for OOXML document generation.
 *
 * @module
 */
import { BaseXmlComponent } from "./base";
import type { Context } from "./base";
import type { IXmlableObject } from "./types";

/**
 * Maps TypeScript property names to their XML attribute names.
 */
export type AttributeMap<T> = Record<keyof T, string>;

/**
 * Simple attribute data as a key-value record.
 */
export type AttributeData = Record<string, boolean | number | string>;

/**
 * Structured attribute payload with explicit key-value mapping.
 */
export type AttributePayload<T> = {
  readonly [P in keyof T]: { readonly key: string; readonly value: T[P] };
};

/**
 * Base class for creating XML attributes with automatic name mapping.
 */
export abstract class XmlAttributeComponent<
  T extends Record<string, any>,
> extends BaseXmlComponent {
  /** Optional mapping from property names to XML attribute names. */
  protected readonly xmlKeys?: AttributeMap<T>;

  public constructor(private readonly root: T) {
    super("_attr");
  }

  public prepForXml(_: Context): IXmlableObject {
    const attrs: Record<string, string> = {};
    Object.entries(this.root).forEach(([key, value]) => {
      if (value !== undefined) {
        const newKey = (this.xmlKeys && this.xmlKeys[key]) || key;
        attrs[newKey] = value;
      }
    });
    return { _attr: attrs };
  }
}

/**
 * Next-generation attribute component with explicit key-value pairs.
 */
export class NextAttributeComponent<T> extends BaseXmlComponent {
  public constructor(private readonly root: AttributePayload<T>) {
    super("_attr");
  }

  public prepForXml(_: Context): IXmlableObject {
    const attrs: Record<string, string | boolean | number> = {};
    const values = Object.values(this.root) as readonly {
      readonly key: string;
      readonly value: string | boolean | number;
    }[];
    for (let i = 0; i < values.length; i++) {
      const { key, value } = values[i];
      if (value !== undefined) {
        attrs[key] = value;
      }
    }
    return { _attr: attrs };
  }

  /** Attribute components produce no XML output — attrs are inlined by the parent. */
  public override toXml(_context: Context): string {
    return "";
  }
}
