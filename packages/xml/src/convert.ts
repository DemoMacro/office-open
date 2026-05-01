import type { Element, Attributes } from "./types";

/**
 * Convert XmlObject (node-xml format) directly to Element (xml-js format).
 * Eliminates the redundant xml() → xml2js() bridge path.
 */
export function toElement(xmlObject: Record<string, any>): Element {
    const tagName = Object.keys(xmlObject)[0];
    const value = xmlObject[tagName];

    const element: Element = {
        type: "element",
        name: tagName,
    };

    if (value == null) {
        return element;
    }

    if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
        element.elements = [{ type: "text", text: String(value) }];
        return element;
    }

    if (Array.isArray(value)) {
        const children: Element[] = [];
        for (const item of value) {
            if (item && typeof item === "object" && "_attr" in item) {
                element.attributes = item._attr as Attributes;
            } else if (item && typeof item === "object") {
                const childKeys = Object.keys(item);
                if (childKeys[0] === "_cdata") {
                    children.push({ type: "cdata", cdata: String(item._cdata) });
                } else {
                    children.push(toElement(item));
                }
            } else if (item != null) {
                children.push({ type: "text", text: String(item) });
            }
        }
        if (children.length > 0) {
            element.elements = children;
        }
        return element;
    }

    if (typeof value === "object") {
        if (value._attr) {
            element.attributes = value._attr as Attributes;
        }
        if (value._cdata) {
            element.elements = [{ type: "cdata", cdata: String(value._cdata) }];
        }
    }

    return element;
}
