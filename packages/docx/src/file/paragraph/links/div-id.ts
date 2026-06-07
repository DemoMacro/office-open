/**
 * Div ID module for WordprocessingML documents.
 *
 * This module provides the div ID element which specifies
 * a reference to an HTML div element.
 *
 * Reference: http://officeopenxml.com/WPparagraph.php
 *
 * @module
 */
import { BuilderElement, numberValObj } from "@file/xml-components";
import type { IXmlableObject, XmlComponent } from "@file/xml-components";

/**
 * Creates a div ID element for a paragraph.
 *
 * The div ID references an HTML div element for HTML conversion purposes.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_DecimalNumber">
 *   <xsd:attribute name="val" type="xsd:integer" use="required"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Set div ID to 1
 * createDivId(1);
 * ```
 */
export const createDivId = (id: number): XmlComponent =>
  new BuilderElement<{ val: number }>({
    attributes: {
      val: { key: "w:val", value: id },
    },
    name: "w:divId",
  });

export const buildDivIdObj = (id: number): IXmlableObject => numberValObj("w:divId", id);
