/**
 * Imported XML Component module for handling external XML content.
 *
 * @module
 */
import { xml2js } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import { XmlAttributeComponent, XmlComponent } from ".";
import type { IContext, IXmlableObject } from "./base";

/**
 * Converts an xml-js Element into an XmlComponent tree.
 */
export const convertToXmlComponent = (
    element: XmlElement,
): ImportedXmlComponent | string | undefined => {
    switch (element.type) {
        case undefined:
        case "element": {
            const xmlComponent = new ImportedXmlComponent(
                element.name as string,
                element.attributes,
            );
            const childElements = element.elements || [];
            for (const childElm of childElements) {
                const child = convertToXmlComponent(childElm);
                if (child !== undefined) {
                    xmlComponent.push(child);
                }
            }
            return xmlComponent;
        }
        case "text": {
            return element.text as string;
        }
        default: {
            return undefined;
        }
    }
};

/**
 * Internal attribute component for imported XML elements.
 * @internal
 */
class ImportedXmlComponentAttributes extends XmlAttributeComponent<any> {}

/**
 * XML component representing imported XML content.
 */
export class ImportedXmlComponent extends XmlComponent {
    public static fromXmlString(importedContent: string): ImportedXmlComponent {
        const xmlObj = xml2js(importedContent, { compact: false }) as XmlElement;
        const root = xmlObj.elements?.[0] ?? xmlObj;
        return convertToXmlComponent(root as XmlElement) as ImportedXmlComponent;
    }

    public constructor(rootKey: string, _attr?: any) {
        super(rootKey);
        if (_attr) {
            this.root.push(new ImportedXmlComponentAttributes(_attr));
        }
    }

    public push(xmlComponent: XmlComponent | string): void {
        this.root.push(xmlComponent);
    }
}

/**
 * Represents attributes for imported root elements.
 */
export class ImportedRootElementAttributes extends XmlComponent {
    public constructor(private readonly _attr: any) {
        super("");
    }

    public prepForXml(_: IContext): IXmlableObject {
        return {
            _attr: this._attr,
        };
    }
}
