/**
 * Math Equation Array module for Office MathML.
 *
 * This module provides the equation array element (m:eqArr) which
 * displays multiple equations stacked vertically.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_eqArr-1.html
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";

import type { MathComponent } from "../math-component";
import { createMathBase } from "../n-ary";
import { createMathEqArrProperties } from "./math-eq-arr-properties";
import type { MathEqArrPropertiesOptions } from "./math-eq-arr-properties";

/**
 * Options for creating a math equation array element.
 *
 * @see {@link MathEqArr}
 */
export interface IMathEqArrOptions {
    /** Equation array properties */
    readonly properties?: MathEqArrPropertiesOptions;
    /** Equations to stack (each row becomes an m:e element) */
    readonly rows: readonly (readonly MathComponent[])[];
}

/**
 * Represents a math equation array in a math equation.
 *
 * The equation array stacks multiple equations vertically with
 * configurable alignment and spacing.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_eqArr-1.html
 *
 * @publicApi
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_EqArr">
 *   <xsd:sequence>
 *     <xsd:element name="eqArrPr" type="CT_EqArrPr" minOccurs="0"/>
 *     <xsd:element name="e" type="CT_OMathArg" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export class MathEqArr extends XmlComponent {
    public constructor(options: IMathEqArrOptions) {
        super("m:eqArr");

        if (options.properties) {
            this.root.push(createMathEqArrProperties(options.properties));
        }

        for (const row of options.rows) {
            this.root.push(createMathBase({ children: row }));
        }
    }
}
