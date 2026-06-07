/**
 * Custom Properties module for WordprocessingML documents.
 *
 * Provides support for managing custom document properties collection.
 *
 * Reference: ISO-IEC29500-4_2016 shared-documentPropertiesCustom.xsd
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";

import { CustomProperty } from "./custom-property";
import type { CustomPropertyOptions } from "./custom-property";

/**
 * Represents the collection of custom document properties.
 *
 * Custom properties allow storing arbitrary metadata as name-value pairs.
 * Each property is assigned a unique ID starting from 2 (per Office specification).
 *
 * Reference: ISO-IEC29500-4_2016 shared-documentPropertiesCustom.xsd
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_CustomProperties">
 *   <xsd:sequence>
 *     <xsd:element name="property" type="CT_Property" minOccurs="0" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * const customProps = new CustomProperties([
 *   { name: "Department", value: "Engineering" },
 *   { name: "Project", value: "Alpha" }
 * ]);
 * ```
 */
export class CustomProperties extends XmlComponent {
  private nextId: number;
  private properties: CustomProperty[] = [];

  public constructor(properties: CustomPropertyOptions[]) {
    super("Properties");

    this.root.push({
      _attr: {
        xmlns: "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
        "xmlns:vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
      },
    });

    // I'm not sure why, but every example I have seen starts with 2
    // https://docs.microsoft.com/en-us/office/open-xml/how-to-set-a-custom-property-in-a-word-processing-document
    this.nextId = 2;

    for (const property of properties) {
      this.addCustomProperty(property);
    }
  }

  public addCustomProperty(property: CustomPropertyOptions): void {
    const prop = new CustomProperty(this.nextId++, property);
    this.properties.push(prop);
    this.root.push(prop);
  }
}
