/**
 * Math Function Properties module for Office MathML.
 *
 * This module provides properties for function structures in math equations.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_funcPr-1.html
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
 * Options for math function properties.
 *
 * @see {@link createMathFunctionProperties}
 */
export interface MathFunctionPropertiesOptions {
    /** Control properties (tracking changes) */
    readonly controlProperties?: MathControlPropertiesOptions;
}

/**
 * Creates properties for a math function structure (m:funcPr).
 *
 * This element specifies properties for the function object,
 * such as function name alignment and spacing.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_funcPr-1.html
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_FuncPr">
 *   <xsd:sequence>
 *     <xsd:element name="ctrlPr" type="CT_CtrlPr" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export const createMathFunctionProperties = (
    options?: MathFunctionPropertiesOptions,
): XmlComponent => {
    const children: XmlComponent[] = [];

    if (options?.controlProperties !== undefined) {
        children.push(createMathControlProperties(options.controlProperties));
    }

    return new BuilderElement({
        children,
        name: "m:funcPr",
    });
};
