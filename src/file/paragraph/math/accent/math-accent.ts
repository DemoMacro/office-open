/**
 * Math Accent module for Office MathML.
 *
 * This module provides the math accent element (e.g., hat, tilde, dot).
 *
 * Reference: ISO/IEC 29500-4, shared-math.xsd, CT_Acc
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

import type { MathComponent } from "../math-component";
import { createMathBase } from "../n-ary";
import { createMathAccentProperties } from "./math-accent-properties";

/**
 * Options for creating a math accent element.
 */
export interface MathAccentOptions {
    /**
     * Accent character (e.g., "\u0302" for hat, "\u0303" for tilde, "\u0307" for dot).
     * If not specified, defaults to hat accent.
     */
    readonly accentCharacter?: string;
    /** Content to place the accent over */
    readonly children: readonly MathComponent[];
}

/**
 * Creates a math accent element (m:acc).
 *
 * An accent places a symbol over (or under) the base expression,
 * such as hat (^), tilde (~), dot (.), arrow (→), etc.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Acc">
 *   <xsd:sequence>
 *     <xsd:element name="accPr" type="CT_AccPr" minOccurs="0"/>
 *     <xsd:element name="e" type="CT_OMathArg"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Hat accent (default)
 * createMathAccent({ children: [new MathRun("x")] });
 * // Tilde accent
 * createMathAccent({ accentCharacter: "\u0303", children: [new MathRun("x")] });
 * // Dot accent
 * createMathAccent({ accentCharacter: "\u0307", children: [new MathRun("x")] });
 * ```
 */
export const createMathAccent = ({ accentCharacter, children }: MathAccentOptions): XmlComponent => {
    const accentProps = accentCharacter ? [createMathAccentProperties({ accentCharacter })] : [];
    return new BuilderElement({
        children: [...accentProps, createMathBase({ children })],
        name: "m:acc",
    });
};
