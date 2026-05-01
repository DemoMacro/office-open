/**
 * Math Limit Location module for Office MathML.
 *
 * This module provides the limit location property for n-ary operators.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_limLoc-1.html
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

/**
 * Options for creating a limit location element.
 */
interface MathLimitLocationOptions {
    /** Location: "undOvr" (under/over) or "subSup" (subscript/superscript). Defaults to "undOvr". */
    readonly value?: string;
}

/**
 * Creates a limit location element for n-ary operators.
 *
 * This element specifies where limits appear relative to the operator:
 * - "undOvr": limits appear directly above and below the operator
 * - "subSup": limits appear as superscript and subscript
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_limLoc-1.html
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_LimLoc">
 *   <xsd:attribute name="val" type="ST_LimLoc"/>
 * </xsd:complexType>
 * ```
 */
export const createMathLimitLocation = ({ value }: MathLimitLocationOptions): XmlComponent =>
    new BuilderElement<Required<MathLimitLocationOptions>>({
        attributes: {
            value: { key: "m:val", value: value || "undOvr" },
        },
        name: "m:limLoc",
    });
