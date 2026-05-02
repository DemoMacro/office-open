/**
 * Math Bar Position module for Office MathML.
 *
 * This module provides the position property for bar structures.
 *
 * Reference: https://www.datypic.com/sc/ooxml/e-m_pos-1.html
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

/**
 * Options for creating a bar position element.
 */
interface MathBarPosOptions {
    /** Position value: "top" for overline, "bot" for underline */
    readonly val: string;
}

/**
 * Creates a position element for bar structures.
 *
 * This element specifies whether the bar appears above (top) or below (bot)
 * the base expression.
 *
 * Reference: https://www.datypic.com/sc/ooxml/e-m_pos-1.html
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_TopBot">
 *   <xsd:attribute name="val" type="ST_TopBot"/>
 * </xsd:complexType>
 * ```
 */
export const createMathBarPos = ({ val }: MathBarPosOptions): XmlComponent =>
    new BuilderElement<MathBarPosOptions>({
        attributes: {
            val: { key: "m:val", value: val },
        },
        name: "m:pos",
    });
