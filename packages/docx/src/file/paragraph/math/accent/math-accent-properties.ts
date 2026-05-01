/**
 * Math Accent Properties module for Office MathML.
 *
 * This module provides properties for the math accent element.
 *
 * Reference: ISO/IEC 29500-4, shared-math.xsd, CT_AccPr
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

/**
 * Options for math accent properties.
 */
export interface MathAccentPropertiesOptions {
    /**
     * Accent character (e.g., "\u0302" for hat, "\u0303" for tilde).
     *
     * Common accent characters:
     * - `\u0302` - circumflex accent (hat): x̂
     * - `\u0303` - tilde: x̃
     * - `\u0307` - dot above: ẋ
     * - `\u0308` - diaeresis (double dot): ẍ
     * - `\u0304` - macron (bar): x̄
     * - `\u0305` - overline: x̅
     * - `\u20D7` - right arrow: x⃗
     * - `\u030C` - caron (háček): x̌
     * - `\u0306` - breve: x̆
     * - `\u030A` - ring above: x̊
     */
    readonly accentCharacter: string;
}

/**
 * Creates math accent properties element (m:accPr).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_AccPr">
 *   <xsd:sequence>
 *     <xsd:element name="chr" type="CT_Char" minOccurs="0"/>
 *     <xsd:element name="ctrlPr" type="CT_CtrlPr" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export const createMathAccentProperties = (options: MathAccentPropertiesOptions): XmlComponent =>
    new BuilderElement({
        children: [
            new BuilderElement<{ readonly val: string }>({
                attributes: {
                    val: { key: "val", value: options.accentCharacter },
                },
                name: "m:chr",
            }),
        ],
        name: "m:accPr",
    });
