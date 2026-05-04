/**
 * Simple XML element types for common OOXML patterns.
 *
 * @module
 */
import { BaseXmlComponent, NextAttributeComponent, XmlComponent } from ".";
import { hpsMeasureValue } from "../values";
import type { PositiveUniversalMeasure } from "../values";
import type { AttributePayload } from "./attributes";
import type { IXmlableObject } from "./types";

// ── Pure functions (zero-allocation) ──

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

// ── Legacy classes (still used by non-hot-path code) ──

/**
 * XML element representing a boolean on/off value (CT_OnOff).
 * @deprecated Use `onOffObj()` for hot-path code.
 */
export class OnOffElement extends XmlComponent {
    public constructor(name: string, val: boolean | undefined = true) {
        super(name);
        if (val !== true) {
            this.root.push(
                new NextAttributeComponent({
                    val: { key: `${name.split(":")[0]}:val`, value: val },
                }),
            );
        }
    }
}

/**
 * XML element representing a half-point size measurement (CT_HpsMeasure).
 * @deprecated Use `hpsMeasureObj()` for hot-path code.
 */
export class HpsMeasureElement extends XmlComponent {
    public constructor(name: string, val: number | PositiveUniversalMeasure) {
        super(name);
        const ns = name.split(":")[0];
        this.root.push(
            new NextAttributeComponent({ val: { key: `${ns}:val`, value: hpsMeasureValue(val) } }),
        );
    }
}

/**
 * XML element representing an empty element (CT_Empty).
 */
export class EmptyElement extends XmlComponent {}

/**
 * XML element with a string value attribute (CT_String).
 * @deprecated Use `stringValObj()` for hot-path code.
 */
export class StringValueElement extends XmlComponent {
    public constructor(name: string, val: string) {
        super(name);
        const ns = name.split(":")[0];
        this.root.push(new NextAttributeComponent({ val: { key: `${ns}:val`, value: val } }));
    }
}

/**
 * XML element with a numeric value attribute.
 * @deprecated Use `numberValObj()` for hot-path code.
 */
export class NumberValueElement extends XmlComponent {
    public constructor(name: string, val: number) {
        super(name);
        const ns = name.split(":")[0];
        this.root.push(new NextAttributeComponent({ val: { key: `${ns}:val`, value: val } }));
    }
}

/**
 * XML element with a string enum value attribute.
 * @deprecated Use `stringEnumValObj()` for hot-path code.
 */
export class StringEnumValueElement<T extends string> extends XmlComponent {
    public constructor(name: string, val: T) {
        super(name);
        const ns = name.split(":")[0];
        this.root.push(new NextAttributeComponent({ val: { key: `${ns}:val`, value: val } }));
    }
}

/**
 * XML element containing text content.
 * @deprecated Use `stringContainerObj()` for hot-path code.
 */
export class StringContainer extends XmlComponent {
    public constructor(name: string, val: string) {
        super(name);
        this.root.push(val);
    }
}

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
        readonly children?: readonly (BaseXmlComponent | string)[];
    }) {
        super(name);

        if (attributes) {
            this.root.push(new NextAttributeComponent(attributes));
        }

        if (children) {
            this.root.push(...children);
        }
    }
}

/**
 * Creates a NextAttributeComponent with explicit XML attribute keys.
 */
export const chartAttr = (attrs: Record<string, string | number | boolean>): BaseXmlComponent =>
    new NextAttributeComponent(
        Object.fromEntries(Object.entries(attrs).map(([key, value]) => [key, { key, value }])),
    );

/**
 * Wraps a component in a named XmlComponent element.
 */
export function wrapEl(elementName: string, child: BaseXmlComponent): XmlComponent {
    const el = new (class extends XmlComponent {
        public constructor(name: string) {
            super(name);
        }
    })(elementName);
    el["root"].push(child);
    return el;
}
