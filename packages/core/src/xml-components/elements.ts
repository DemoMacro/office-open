/**
 * Simple XML element types for common OOXML patterns.
 *
 * @module
 */
import { BaseXmlComponent, NextAttributeComponent, XmlComponent } from ".";
import { hpsMeasureValue } from "../values";
import type { PositiveUniversalMeasure } from "../values";
import type { AttributePayload } from "./attributes";

/**
 * XML element representing a boolean on/off value (CT_OnOff).
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
        readonly children?: readonly XmlComponent[];
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
