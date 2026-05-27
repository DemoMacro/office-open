/**
 * Imported XML Component module for handling external XML content.
 *
 * @module
 */
import { xml2js } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import { XmlComponent } from ".";
import type { Context, IXmlableObject } from "./base";

/**
 * Converts an xml-js Element into an XmlComponent tree.
 */
export const convertToXmlComponent = (
  element: XmlElement,
): ImportedXmlComponent | string | undefined => {
  switch (element.type) {
    case undefined:
    case "element": {
      const xmlComponent = new ImportedXmlComponent(element.name as string, element.attributes);
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
 * XML component representing imported XML content.
 */
export class ImportedXmlComponent extends XmlComponent {
  protected _sourceXml?: string;

  public static fromXmlString(importedContent: string): ImportedXmlComponent {
    const xmlObj = xml2js(importedContent, { compact: false }) as XmlElement;
    const root = xmlObj.elements?.[0] ?? xmlObj;
    const component = convertToXmlComponent(root as XmlElement) as ImportedXmlComponent;
    component._sourceXml = importedContent;
    return component;
  }

  public get sourceXml(): string | undefined {
    return this._sourceXml;
  }

  public override toXml(_context: Context): string {
    return this._sourceXml ?? super.toXml(_context);
  }

  public constructor(rootKey: string, _attr?: any) {
    super(rootKey);
    if (_attr) {
      this.root.push({ _attr });
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

  public prepForXml(_: Context): IXmlableObject {
    return {
      _attr: this._attr,
    };
  }
}
