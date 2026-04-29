/**
 * Math Equation Array Properties module for Office MathML.
 *
 * This module provides properties for the equation array element (CT_EqArrPr).
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_eqArrPr-1.html
 *
 * @module
 */
import { BuilderElement, OnOffElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

/**
 * Options for math equation array properties.
 *
 * @see {@link createMathEqArrProperties}
 */
export interface MathEqArrPropertiesOptions {
    /** Base justification (vertical alignment) */
    readonly baseJc?: "top" | "bot" | "center";
    /** Maximum distance between equations */
    readonly maxDist?: boolean;
    /** Object distance */
    readonly objDist?: boolean;
    /** Row spacing rule */
    readonly rSpRule?: "single" | "1.5" | "double" | "exactly" | "multiple";
    /** Row spacing value */
    readonly rSp?: number;
}

/**
 * Creates math equation array properties element (m:eqArrPr).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_EqArrPr">
 *   <xsd:sequence>
 *     <xsd:element name="baseJc" type="CT_YAlign" minOccurs="0"/>
 *     <xsd:element name="maxDist" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="objDist" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="rSpRule" type="CT_SpacingRule" minOccurs="0"/>
 *     <xsd:element name="rSp" type="CT_UnSignedInteger" minOccurs="0"/>
 *     <xsd:element name="ctrlPr" type="CT_CtrlPr" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export const createMathEqArrProperties = (options: MathEqArrPropertiesOptions): XmlComponent => {
    const children: XmlComponent[] = [];

    if (options.baseJc !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: options.baseJc } },
                name: "m:baseJc",
            }),
        );
    }
    if (options.maxDist !== undefined) {
        children.push(new OnOffElement("m:maxDist", options.maxDist));
    }
    if (options.objDist !== undefined) {
        children.push(new OnOffElement("m:objDist", options.objDist));
    }
    if (options.rSpRule !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: options.rSpRule } },
                name: "m:rSpRule",
            }),
        );
    }
    if (options.rSp !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: options.rSp } },
                name: "m:rSp",
            }),
        );
    }

    return new BuilderElement({
        children,
        name: "m:eqArrPr",
    });
};
