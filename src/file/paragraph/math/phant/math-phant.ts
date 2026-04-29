/**
 * Math Phantom module for Office MathML.
 *
 * This module provides the phantom element (m:phant) which makes
 * content invisible while preserving spacing.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_phant-1.html
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";

import type { MathComponent } from "../math-component";
import { createMathBase } from "../n-ary";
import { createMathPhantProperties } from "./math-phant-properties";
import type { MathPhantPropertiesOptions } from "./math-phant-properties";

/**
 * Options for creating a math phantom element.
 *
 * @see {@link MathPhant}
 */
export interface IMathPhantOptions {
    /** Phantom properties */
    readonly properties?: MathPhantPropertiesOptions;
    /** Content to be made invisible */
    readonly children: readonly MathComponent[];
}

/**
 * Represents a math phantom in a math equation.
 *
 * The phantom element makes its content invisible while preserving
 * the spacing the content would normally occupy.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_phant-1.html
 *
 * @publicApi
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Phant">
 *   <xsd:sequence>
 *     <xsd:element name="phantPr" type="CT_PhantPr" minOccurs="0"/>
 *     <xsd:element name="e" type="CT_OMathArg"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export class MathPhant extends XmlComponent {
    public constructor(options: IMathPhantOptions) {
        super("m:phant");

        if (options.properties) {
            this.root.push(createMathPhantProperties(options.properties));
        }

        this.root.push(createMathBase({ children: options.children }));
    }
}
