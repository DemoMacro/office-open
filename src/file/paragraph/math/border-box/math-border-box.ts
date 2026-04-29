/**
 * Math Border Box module for Office MathML.
 *
 * This module provides the math border box element (m:borderBox) which
 * draws borders around math expressions.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_borderBox-1.html
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";

import type { MathComponent } from "../math-component";
import { createMathBase } from "../n-ary";
import { createMathBorderBoxProperties } from "./math-border-box-properties";
import type { MathBorderBoxPropertiesOptions } from "./math-border-box-properties";

/**
 * Options for creating a math border box element.
 *
 * @see {@link MathBorderBox}
 */
export interface IMathBorderBoxOptions {
    /** Border box properties */
    readonly properties?: MathBorderBoxPropertiesOptions;
    /** Content to be bordered */
    readonly children: readonly MathComponent[];
}

/**
 * Represents a math border box in a math equation.
 *
 * The border box element draws visible borders around math expressions,
 * with options to hide specific sides or add strike-through lines.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_borderBox-1.html
 *
 * @publicApi
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_BorderBox">
 *   <xsd:sequence>
 *     <xsd:element name="borderBoxPr" type="CT_BorderBoxPr" minOccurs="0"/>
 *     <xsd:element name="e" type="CT_OMathArg"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export class MathBorderBox extends XmlComponent {
    public constructor(options: IMathBorderBoxOptions) {
        super("m:borderBox");

        if (options.properties) {
            this.root.push(createMathBorderBoxProperties(options.properties));
        }

        this.root.push(createMathBase({ children: options.children }));
    }
}
