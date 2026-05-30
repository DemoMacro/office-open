/**
 * Simple XML element types for common OOXML patterns.
 *
 * @module
 */
import { BaseXmlComponent, XmlComponent } from ".";
import { hpsMeasureValue } from "../values";
import type { PositiveUniversalMeasure } from "../values";
import type { AttributePayload } from "./attributes";
import { EMPTY_OBJECT } from "./component";
import type { IXmlableObject } from "./types";

// ── Pure functions (zero-allocation) ──

/**
 * Build an XML element with arbitrary attributes, filtering out undefined values.
 */
export function attrObj(
  name: string,
  attrs: Record<string, string | number | boolean | undefined>,
): IXmlableObject {
  const filtered: Record<string, string | number | boolean> = {};
  for (const [key, val] of Object.entries(attrs)) {
    if (val !== undefined) filtered[key] = val;
  }
  return { [name]: { _attr: filtered } };
}

const ON_OFF_TRUE_CACHE = new Map<string, IXmlableObject>();

/**
 * Build a CT_OnOff XML object without allocating any XmlComponent.
 * `val=true` returns a frozen singleton (cached per name).
 */
export function onOffObj(name: string, val: boolean | undefined = true): IXmlableObject {
  if (val === true) {
    let cached = ON_OFF_TRUE_CACHE.get(name);
    if (!cached) {
      cached = Object.freeze({ [name]: Object.freeze({}) });
      ON_OFF_TRUE_CACHE.set(name, cached);
    }
    return cached;
  }
  const ns = name.split(":")[0];
  return { [name]: { _attr: { [`${ns}:val`]: val } } };
}

/**
 * Build a CT_HpsMeasure XML object (half-point size) without allocation.
 */
export function hpsMeasureObj(
  name: string,
  val: number | PositiveUniversalMeasure,
): IXmlableObject {
  const ns = name.split(":")[0];
  return { [name]: { _attr: { [`${ns}:val`]: hpsMeasureValue(val) } } };
}

/**
 * Build a CT_String XML object (string value attribute) without allocation.
 */
export function stringValObj(name: string, val: string): IXmlableObject {
  const ns = name.split(":")[0];
  return { [name]: { _attr: { [`${ns}:val`]: val } } };
}

/**
 * Build a numeric value attribute XML object without allocation.
 */
export function numberValObj(name: string, val: number): IXmlableObject {
  const ns = name.split(":")[0];
  return { [name]: { _attr: { [`${ns}:val`]: val } } };
}

/**
 * Build a string enum value attribute XML object without allocation.
 */
export function stringEnumValObj<T extends string>(name: string, val: T): IXmlableObject {
  const ns = name.split(":")[0];
  return { [name]: { _attr: { [`${ns}:val`]: val } } };
}

/**
 * Build an element wrapping a text string without allocation.
 */
export function stringContainerObj(name: string, val: string): IXmlableObject {
  return { [name]: [val] };
}

// ── Class-based elements ──

/**
 * XML element representing an empty element (CT_Empty).
 */
export class EmptyElement extends XmlComponent {}

/**
 * Flexible XML element builder with explicit attribute and child configuration.
 */
export class BuilderElement<T = {}> extends XmlComponent {
  public constructor({
    name,
    attributes,
    children,
  }: {
    readonly name: string;
    readonly attributes?: AttributePayload<T>;
    readonly children?: readonly (BaseXmlComponent | IXmlableObject | string)[];
  }) {
    super(name);

    if (attributes) {
      // Build _attr object directly instead of creating NextAttributeComponent.
      // This saves one class allocation and one toXml() call per element.
      const attrs: Record<string, string | number | boolean> = {};
      const vals = Object.values(attributes) as readonly {
        readonly key: string;
        readonly value: string | number | boolean;
      }[];
      for (let i = 0; i < vals.length; i++) {
        const { key, value } = vals[i];
        if (value !== undefined) {
          attrs[key] = value;
        }
      }
      this.root.push({ _attr: attrs });
    }

    if (children) {
      this.root.push(...children);
    }
  }
}

/**
 * Builds an attribute-only IXmlableObject from a plain key-value record.
 */
export function buildAttrObject(attrs: Record<string, string | number | boolean>): IXmlableObject {
  const result: Record<string, string | number | boolean> = {};
  for (const [key, value] of Object.entries(attrs)) {
    if (value !== undefined) result[key] = value;
  }
  return Object.keys(result).length > 0 ? { _attr: result } : EMPTY_OBJECT;
}

/**
 * Creates an attribute IXmlableObject for chart elements.
 */
export const chartAttr = (attrs: Record<string, string | number | boolean>): IXmlableObject =>
  buildAttrObject(attrs);

/**
 * Wraps a component in a named XmlComponent element.
 */
export function wrapEl(
  elementName: string,
  child: BaseXmlComponent | IXmlableObject,
): XmlComponent {
  const el = new (class extends XmlComponent {
    public constructor(name: string) {
      super(name);
    }
  })(elementName);
  el["root"].push(child);
  return el;
}
