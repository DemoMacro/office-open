/**
 * Math Bracket Properties module for Office MathML.
 *
 * This module provides properties for bracket delimiters in math equations.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_dPr-1.html
 *
 * @module
 */
import { BuilderElement, OnOffElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

import { createMathBeginningCharacter } from "./math-beginning-character";
import { createMathEndingCharacter } from "./math-ending-char";

/**
 * Options for creating math bracket properties.
 *
 * @see {@link createMathBracketProperties}
 */
export interface MathBracketPropertiesOptions {
    /** Optional custom characters for bracket delimiters */
    readonly characters?: {
        /** The opening/beginning bracket character */
        readonly beginningCharacter: string;
        /** The closing/ending bracket character */
        readonly endingCharacter: string;
    };
    /** Separator character for multiple arguments */
    readonly separatorCharacter?: string;
    /** Whether the delimiters grow to match their content */
    readonly grow?: boolean;
    /** Delimiter shape: "centered" or "match" */
    readonly shape?: "centered" | "match";
}

/**
 * Creates bracket properties for math delimiter objects.
 *
 * This element specifies properties for the delimiter object,
 * including custom beginning and ending characters.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_dPr-1.html
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_DPr">
 *   <xsd:sequence>
 *     <xsd:element name="begChr" type="CT_Char" minOccurs="0"/>
 *     <xsd:element name="sepChr" type="CT_Char" minOccurs="0"/>
 *     <xsd:element name="endChr" type="CT_Char" minOccurs="0"/>
 *     <xsd:element name="grow" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="shp" type="CT_Shp" minOccurs="0"/>
 *     <xsd:element name="ctrlPr" type="CT_CtrlPr" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export const createMathBracketProperties = ({
    characters,
    separatorCharacter,
    grow,
    shape,
}: MathBracketPropertiesOptions): XmlComponent => {
    const children: XmlComponent[] = [];

    if (characters) {
        children.push(createMathBeginningCharacter({ character: characters.beginningCharacter }));
    }
    if (separatorCharacter !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: separatorCharacter } },
                name: "m:sepChr",
            }),
        );
    }
    if (characters) {
        children.push(createMathEndingCharacter({ character: characters.endingCharacter }));
    }
    if (grow !== undefined) {
        children.push(new OnOffElement("m:grow", grow));
    }
    if (shape !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: shape } },
                name: "m:shp",
            }),
        );
    }

    return new BuilderElement({
        children,
        name: "m:dPr",
    });
};
