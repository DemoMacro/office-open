/**
 * Math Group Character module for Office MathML.
 *
 * This module provides the group character element (m:groupChr) which
 * places a character (like brace or bracket) over or under an expression.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_groupChr-1.html
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";

import type { MathComponent } from "../math-component";
import { createMathBase } from "../n-ary";
import { createMathGroupChrProperties } from "./math-group-chr-properties";
import type { MathGroupChrPropertiesOptions } from "./math-group-chr-properties";

/**
 * Options for creating a math group character element.
 *
 * @see {@link MathGroupChr}
 */
export interface IMathGroupChrOptions {
    /** Group character properties */
    readonly properties?: MathGroupChrPropertiesOptions;
    /** Content under/over the group character */
    readonly children: readonly MathComponent[];
}

/**
 * Represents a math group character in a math equation.
 *
 * The group character places a character (such as a brace, bracket,
 * or arrow) over or under a math expression.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_groupChr-1.html
 *
 * @publicApi
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_GroupChr">
 *   <xsd:sequence>
 *     <xsd:element name="groupChrPr" type="CT_GroupChrPr" minOccurs="0"/>
 *     <xsd:element name="e" type="CT_OMathArg"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export class MathGroupChr extends XmlComponent {
    public constructor(options: IMathGroupChrOptions) {
        super("m:groupChr");

        if (options.properties) {
            this.root.push(createMathGroupChrProperties(options.properties));
        }

        this.root.push(createMathBase({ children: options.children }));
    }
}
