/**
 * Math Border Box Properties module for Office MathML.
 *
 * This module provides properties for the math border box element (CT_BorderBoxPr).
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_borderBoxPr-1.html
 *
 * @module
 */
import { BuilderElement, OnOffElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

/**
 * Options for math border box properties.
 *
 * @see {@link createMathBorderBoxProperties}
 */
export interface MathBorderBoxPropertiesOptions {
    /** Hide top border */
    readonly hideTop?: boolean;
    /** Hide bottom border */
    readonly hideBot?: boolean;
    /** Hide left border */
    readonly hideLeft?: boolean;
    /** Hide right border */
    readonly hideRight?: boolean;
    /** Horizontal strike-through */
    readonly strikeH?: boolean;
    /** Vertical strike-through */
    readonly strikeV?: boolean;
    /** Bottom-left to top-right diagonal strike */
    readonly strikeBLTR?: boolean;
    /** Top-left to bottom-right diagonal strike */
    readonly strikeTLBR?: boolean;
}

/**
 * Creates math border box properties element (m:borderBoxPr).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_BorderBoxPr">
 *   <xsd:sequence>
 *     <xsd:element name="hideTop" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="hideBot" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="hideLeft" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="hideRight" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="strikeH" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="strikeV" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="strikeBLTR" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="strikeTLBR" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="ctrlPr" type="CT_CtrlPr" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export const createMathBorderBoxProperties = (
    options: MathBorderBoxPropertiesOptions,
): XmlComponent => {
    const children: XmlComponent[] = [];

    if (options.hideTop !== undefined) {
        children.push(new OnOffElement("m:hideTop", options.hideTop));
    }
    if (options.hideBot !== undefined) {
        children.push(new OnOffElement("m:hideBot", options.hideBot));
    }
    if (options.hideLeft !== undefined) {
        children.push(new OnOffElement("m:hideLeft", options.hideLeft));
    }
    if (options.hideRight !== undefined) {
        children.push(new OnOffElement("m:hideRight", options.hideRight));
    }
    if (options.strikeH !== undefined) {
        children.push(new OnOffElement("m:strikeH", options.strikeH));
    }
    if (options.strikeV !== undefined) {
        children.push(new OnOffElement("m:strikeV", options.strikeV));
    }
    if (options.strikeBLTR !== undefined) {
        children.push(new OnOffElement("m:strikeBLTR", options.strikeBLTR));
    }
    if (options.strikeTLBR !== undefined) {
        children.push(new OnOffElement("m:strikeTLBR", options.strikeTLBR));
    }

    return new BuilderElement({
        children,
        name: "m:borderBoxPr",
    });
};
