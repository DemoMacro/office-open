/**
 * Math Matrix module for Office MathML.
 *
 * This module provides the matrix element (m:m) for displaying
 * mathematical matrices.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_m-1.html
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import { XmlComponent } from "@file/xml-components";

import type { MathComponent } from "../math-component";
import { createMathBase } from "../n-ary";
import { createMathMatrixProperties } from "./math-matrix-properties";
import type { MathMatrixPropertiesOptions } from "./math-matrix-properties";

/**
 * Options for creating a math matrix element.
 *
 * @see {@link MathMatrix}
 */
export interface IMathMatrixOptions {
    /** Matrix properties */
    readonly properties?: MathMatrixPropertiesOptions;
    /** Matrix rows (each row is an array of math components) */
    readonly rows: readonly (readonly MathComponent[])[];
}

/**
 * Represents a math matrix in a math equation.
 *
 * The matrix element displays a grid of math expressions arranged
 * in rows and columns.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_m-1.html
 *
 * @publicApi
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_M">
 *   <xsd:sequence>
 *     <xsd:element name="mPr" type="CT_MPr" minOccurs="0"/>
 *     <xsd:element name="mr" type="CT_MR" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export class MathMatrix extends XmlComponent {
    public constructor(options: IMathMatrixOptions) {
        super("m:m");

        if (options.properties) {
            this.root.push(createMathMatrixProperties(options.properties));
        }

        for (const row of options.rows) {
            this.root.push(
                new BuilderElement({
                    children: row.map((cell) =>
                        createMathBase({ children: Array.isArray(cell) ? cell : [cell] }),
                    ),
                    name: "m:mr",
                }),
            );
        }
    }
}
