/**
 * Math Lower Limit Properties module for Office MathML.
 *
 * This module provides properties for lower limit structures (CT_LimLowPr).
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_limLowPr-1.html
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
 * Options for math lower limit properties.
 *
 * @see {@link createMathLimitLowProperties}
 */
export interface MathLimitLowPropertiesOptions {
    /** Control properties (tracking changes) */
    readonly controlProperties?: MathControlPropertiesOptions;
}

/**
 * Creates properties for a lower limit structure (m:limLowPr).
 *
 * This element specifies properties for the lower limit object.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_limLowPr-1.html
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_LimLowPr">
 *   <xsd:sequence>
 *     <xsd:element name="ctrlPr" type="CT_CtrlPr" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export const createMathLimitLowProperties = (
    options?: MathLimitLowPropertiesOptions,
): XmlComponent => {
    const children: XmlComponent[] = [];

    if (options?.controlProperties !== undefined) {
        children.push(createMathControlProperties(options.controlProperties));
    }

    return new BuilderElement({
        children,
        name: "m:limLowPr",
    });
};
