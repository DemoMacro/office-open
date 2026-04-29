/**
 * Math Sub-Super-Script Properties module for Office MathML.
 *
 * This module provides properties for combined subscript and superscript structures.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_sSubSupPr-1.html
 *
 * @module
 */
import { BuilderElement, OnOffElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

/**
 * Options for math sub-super-script properties.
 *
 * @see {@link createMathSubSuperScriptProperties}
 */
export interface MathSubSuperScriptPropertiesOptions {
    /** Align scripts to the same vertical position */
    readonly alignScripts?: boolean;
}

/**
 * Creates properties for a combined subscript and superscript structure.
 *
 * This element specifies properties for the subscript-superscript object.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_sSubSupPr-1.html
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_SSubSupPr">
 *   <xsd:sequence>
 *     <xsd:element name="alnScr" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="ctrlPr" type="CT_CtrlPr" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export const createMathSubSuperScriptProperties = (
    options?: MathSubSuperScriptPropertiesOptions,
): XmlComponent => {
    const children: XmlComponent[] = [];

    if (options?.alignScripts !== undefined) {
        children.push(new OnOffElement("m:alnScr", options.alignScripts));
    }

    return new BuilderElement({
        children,
        name: "m:sSubSupPr",
    });
};
