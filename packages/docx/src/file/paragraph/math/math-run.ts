/**
 * Math Run module for Office MathML.
 *
 * This module provides the MathRun class for text content within math equations.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_r-1.html
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";

import { createMathRunProperties, type MathRunPropertiesOptions } from "./math-run-properties";
import { MathText } from "./math-text";

/**
 * Options for creating a math run.
 *
 * @see {@link MathRun}
 */
export interface MathRunOptions {
    /** The text content */
    readonly text: string;
    /** Optional run properties */
    readonly properties?: MathRunPropertiesOptions;
}

/**
 * Represents a run of text within a math equation.
 *
 * MathRun is the container for text content in Office MathML,
 * similar to how Run contains text in regular paragraphs.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_r-1.html
 *
 * @publicApi
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_R">
 *   <xsd:sequence>
 *     <xsd:element name="rPr" type="CT_RPR" minOccurs="0"/>
 *     <xsd:group ref="EG_ScriptStyle" minOccurs="0"/>
 *     <xsd:choice minOccurs="0" maxOccurs="unbounded">
 *       <xsd:element ref="w:br"/>
 *       <xsd:element name="t" type="CT_Text"/>
 *     </xsd:choice>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Simple text
 * new MathRun("x + y");
 *
 * // With properties
 * new MathRun({ text: "x", properties: { script: "fraktur", style: "b" } });
 * ```
 */
export class MathRun extends XmlComponent {
    public constructor(textOrOptions: string | MathRunOptions) {
        super("m:r");

        const options = typeof textOrOptions === "string" ? { text: textOrOptions } : textOrOptions;

        if (options.properties) {
            this.root.push(createMathRunProperties(options.properties));
        }

        this.root.push(new MathText(options.text));
    }
}
