/**
 * Math Matrix Properties module for Office MathML.
 *
 * This module provides properties for the matrix element (CT_MPr).
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_mPr-1.html
 *
 * @module
 */
import { BuilderElement, OnOffElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

/**
 * Options for math matrix properties.
 *
 * @see {@link createMathMatrixProperties}
 */
export interface MathMatrixPropertiesOptions {
    /** Base justification (vertical alignment of cells) */
    readonly baseJc?: "top" | "bot" | "center";
    /** Hide place holders */
    readonly plcHide?: boolean;
    /** Row spacing rule */
    readonly rSpRule?: "single" | "1.5" | "double" | "exactly" | "multiple";
    /** Column gap rule */
    readonly cGpRule?: "single" | "1.5" | "double" | "exactly" | "multiple";
    /** Row spacing */
    readonly rSp?: number;
    /** Column spacing */
    readonly cSp?: number;
    /** Column group spacing */
    readonly cGp?: number;
}

/**
 * Creates math matrix properties element (m:mPr).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_MPr">
 *   <xsd:sequence>
 *     <xsd:element name="baseJc" type="CT_YAlign" minOccurs="0"/>
 *     <xsd:element name="plcHide" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="rSpRule" type="CT_SpacingRule" minOccurs="0"/>
 *     <xsd:element name="cGpRule" type="CT_SpacingRule" minOccurs="0"/>
 *     <xsd:element name="rSp" type="CT_UnSignedInteger" minOccurs="0"/>
 *     <xsd:element name="cSp" type="CT_UnSignedInteger" minOccurs="0"/>
 *     <xsd:element name="cGp" type="CT_UnSignedInteger" minOccurs="0"/>
 *     <xsd:element name="mcs" type="CT_MCS" minOccurs="0"/>
 *     <xsd:element name="ctrlPr" type="CT_CtrlPr" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export const createMathMatrixProperties = (options: MathMatrixPropertiesOptions): XmlComponent => {
    const children: XmlComponent[] = [];

    if (options.baseJc !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: options.baseJc } },
                name: "m:baseJc",
            }),
        );
    }
    if (options.plcHide !== undefined) {
        children.push(new OnOffElement("m:plcHide", options.plcHide));
    }
    if (options.rSpRule !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: options.rSpRule } },
                name: "m:rSpRule",
            }),
        );
    }
    if (options.cGpRule !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: options.cGpRule } },
                name: "m:cGpRule",
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
    if (options.cSp !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: options.cSp } },
                name: "m:cSp",
            }),
        );
    }
    if (options.cGp !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: options.cGp } },
                name: "m:cGp",
            }),
        );
    }

    return new BuilderElement({
        children,
        name: "m:mPr",
    });
};
