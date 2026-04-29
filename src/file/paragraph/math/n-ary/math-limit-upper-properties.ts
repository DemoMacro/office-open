/**
 * Math Upper Limit Properties module for Office MathML.
 *
 * This module provides properties for upper limit structures (CT_LimUppPr).
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_limUppPr-1.html
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

import {
    createMathControlProperties,
    type MathControlPropertiesOptions,
} from "../math-control-properties";

/**
 * Options for math upper limit properties.
 *
 * @see {@link createMathLimitUpperProperties}
 */
export interface MathLimitUpperPropertiesOptions {
    /** Control properties (tracking changes) */
    readonly controlProperties?: MathControlPropertiesOptions;
}

/**
 * Creates properties for an upper limit structure (m:limUppPr).
 *
 * This element specifies properties for the upper limit object.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_limUppPr-1.html
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_LimUppPr">
 *   <xsd:sequence>
 *     <xsd:element name="ctrlPr" type="CT_CtrlPr" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export const createMathLimitUpperProperties = (
    options?: MathLimitUpperPropertiesOptions,
): XmlComponent => {
    const children: XmlComponent[] = [];

    if (options?.controlProperties !== undefined) {
        children.push(createMathControlProperties(options.controlProperties));
    }

    return new BuilderElement({
        children,
        name: "m:limUppPr",
    });
};
