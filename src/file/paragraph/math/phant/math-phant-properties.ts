/**
 * Math Phantom Properties module for Office MathML.
 *
 * This module provides properties for the phantom element (CT_PhantPr).
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_phantPr-1.html
 *
 * @module
 */
import { BuilderElement, OnOffElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

/**
 * Options for math phantom properties.
 *
 * @see {@link createMathPhantProperties}
 */
export interface MathPhantPropertiesOptions {
    /** Show the phantom content */
    readonly show?: boolean;
    /** Zero width */
    readonly zeroWid?: boolean;
    /** Zero ascent */
    readonly zeroAsc?: boolean;
    /** Zero descent */
    readonly zeroDesc?: boolean;
    /** Transparent */
    readonly transp?: boolean;
}

/**
 * Creates math phantom properties element (m:phantPr).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_PhantPr">
 *   <xsd:sequence>
 *     <xsd:element name="show" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="zeroWid" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="zeroAsc" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="zeroDesc" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="transp" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="ctrlPr" type="CT_CtrlPr" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export const createMathPhantProperties = (options: MathPhantPropertiesOptions): XmlComponent => {
    const children: XmlComponent[] = [];

    if (options.show !== undefined) {
        children.push(new OnOffElement("m:show", options.show));
    }
    if (options.zeroWid !== undefined) {
        children.push(new OnOffElement("m:zeroWid", options.zeroWid));
    }
    if (options.zeroAsc !== undefined) {
        children.push(new OnOffElement("m:zeroAsc", options.zeroAsc));
    }
    if (options.zeroDesc !== undefined) {
        children.push(new OnOffElement("m:zeroDesc", options.zeroDesc));
    }
    if (options.transp !== undefined) {
        children.push(new OnOffElement("m:transp", options.transp));
    }

    return new BuilderElement({
        children,
        name: "m:phantPr",
    });
};
