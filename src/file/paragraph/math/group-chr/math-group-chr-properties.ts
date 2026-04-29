/**
 * Math Group Character Properties module for Office MathML.
 *
 * This module provides properties for the group character element (CT_GroupChrPr).
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_groupChrPr-1.html
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

/**
 * Options for math group character properties.
 *
 * @see {@link createMathGroupChrProperties}
 */
export interface MathGroupChrPropertiesOptions {
    /** Character to display (e.g., "\u23DF" for bottom curly bracket) */
    readonly chr?: string;
    /** Position of the character: "top" or "bot" */
    readonly pos?: "top" | "bot";
    /** Vertical justification: "top" or "bot" */
    readonly vertJc?: "top" | "bot";
}

/**
 * Creates math group character properties element (m:groupChrPr).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_GroupChrPr">
 *   <xsd:sequence>
 *     <xsd:element name="chr" type="CT_Char" minOccurs="0"/>
 *     <xsd:element name="pos" type="CT_TopBot" minOccurs="0"/>
 *     <xsd:element name="vertJc" type="CT_TopBot" minOccurs="0"/>
 *     <xsd:element name="ctrlPr" type="CT_CtrlPr" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export const createMathGroupChrProperties = (
    options: MathGroupChrPropertiesOptions,
): XmlComponent => {
    const children: XmlComponent[] = [];

    if (options.chr !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: options.chr } },
                name: "m:chr",
            }),
        );
    }
    if (options.pos !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: options.pos } },
                name: "m:pos",
            }),
        );
    }
    if (options.vertJc !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: options.vertJc } },
                name: "m:vertJc",
            }),
        );
    }

    return new BuilderElement({
        children,
        name: "m:groupChrPr",
    });
};
