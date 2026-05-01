/**
 * Math Box module for Office MathML.
 *
 * This module provides the math box element (m:box) which groups
 * math expressions as a single entity.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_box-1.html
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";

import type { MathComponent } from "../math-component";
import { createMathBase } from "../n-ary";
import { createMathBoxProperties } from "./math-box-properties";
import type { MathBoxPropertiesOptions } from "./math-box-properties";

/**
 * Options for creating a math box element.
 *
 * @see {@link MathBox}
 */
export interface IMathBoxOptions {
    /** Box properties */
    readonly properties?: MathBoxPropertiesOptions;
    /** Content to be boxed */
    readonly children: readonly MathComponent[];
}

/**
 * Represents a math box in a math equation.
 *
 * The box element groups math expressions as a single entity,
 * controlling line breaking, alignment, and other grouping behavior.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_box-1.html
 *
 * @publicApi
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Box">
 *   <xsd:sequence>
 *     <xsd:element name="boxPr" type="CT_BoxPr" minOccurs="0"/>
 *     <xsd:element name="e" type="CT_OMathArg"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export class MathBox extends XmlComponent {
    public constructor(options: IMathBoxOptions) {
        super("m:box");

        if (options.properties) {
            this.root.push(createMathBoxProperties(options.properties));
        }

        this.root.push(createMathBase({ children: options.children }));
    }
}
