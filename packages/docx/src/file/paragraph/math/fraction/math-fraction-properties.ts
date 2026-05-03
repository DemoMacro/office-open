/**
 * Math Fraction Properties module for Office MathML.
 *
 * This module provides properties for the math fraction element.
 *
 * Reference: ISO/IEC 29500-4, shared-math.xsd, CT_FPr
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

/**
 * Fraction type (how the fraction bar is displayed).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_FType">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="bar"/>
 *     <xsd:enumeration value="skw"/>
 *     <xsd:enumeration value="lin"/>
 *     <xsd:enumeration value="noBar"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 *
 * @publicApi
 */
export const FractionType = {
    /** Standard fraction with horizontal bar: a/b */
    BAR: "bar",
    /** Skewed fraction: a/b (diagonal) */
    SKEWED: "skw",
    /** Linear fraction: a/b (horizontal, no bar) */
    LINEAR: "lin",
    /** No bar fraction: a b (stacked, no bar) */
    NO_BAR: "noBar",
} as const;

/**
 * Options for math fraction properties.
 */
export interface MathFractionPropertiesOptions {
    /** Fraction type (bar, skewed, linear, or no bar) */
    readonly fractionType?: keyof typeof FractionType;
}

/**
 * Creates math fraction properties element (m:fPr).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_FPr">
 *   <xsd:sequence>
 *     <xsd:element name="type" type="CT_FType" minOccurs="0"/>
 *     <xsd:element name="ctrlPr" type="CT_CtrlPr" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export const createMathFractionProperties = (
    options: MathFractionPropertiesOptions,
): XmlComponent => {
    const children: XmlComponent[] = [];

    if (options.fractionType) {
        children.push(
            new BuilderElement<{ readonly val: string }>({
                attributes: {
                    val: { key: "m:val", value: FractionType[options.fractionType] },
                },
                name: "m:type",
            }),
        );
    }

    return new BuilderElement({
        children,
        name: "m:fPr",
    });
};
