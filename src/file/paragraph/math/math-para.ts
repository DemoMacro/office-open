/**
 * Office Math Paragraph module for WordprocessingML documents.
 *
 * This module provides the math paragraph element (m:oMathPara) which
 * is a container for one or more math equations with paragraph-level
 * formatting.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_oMathPara-1.html
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import { XmlComponent } from "@file/xml-components";

import { Math } from "./math";
import type { IMathOptions } from "./math";

/**
 * Math paragraph justification types.
 */
export type MathJustification = "left" | "right" | "center" | "centerGroup";

/**
 * Options for creating a MathParagraph element.
 *
 * @see {@link MathParagraph}
 */
export interface IMathParagraphOptions {
    /** Justification for the math paragraph */
    readonly justification?: MathJustification;
    /** Math equations in this paragraph */
    readonly children: readonly IMathOptions[];
}

/**
 * Creates math paragraph properties element (m:oMathParaPr).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_OMathParaPr">
 *   <xsd:sequence>
 *     <xsd:element name="jc" type="CT_OMathJc" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
const createMathParagraphProperties = (justification: MathJustification): XmlComponent =>
    new BuilderElement({
        children: [
            new BuilderElement({
                attributes: { val: { key: "m:val", value: justification } },
                name: "m:jc",
            }),
        ],
        name: "m:oMathParaPr",
    });

/**
 * Represents a math paragraph in a WordprocessingML document.
 *
 * MathParagraph is a container for one or more math equations with
 * paragraph-level alignment. It groups math equations and applies
 * consistent justification.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_oMathPara-1.html
 *
 * @publicApi
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_OMathPara">
 *   <xsd:sequence>
 *     <xsd:element name="oMathParaPr" type="CT_OMathParaPr" minOccurs="0"/>
 *     <xsd:element name="oMath" type="CT_OMath" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * new MathParagraph({
 *   justification: "center",
 *   children: [
 *     { children: [new MathRun("x + y = z")] },
 *   ],
 * });
 * ```
 */
export class MathParagraph extends XmlComponent {
    public constructor(options: IMathParagraphOptions) {
        super("m:oMathPara");

        if (options.justification) {
            this.root.push(createMathParagraphProperties(options.justification));
        }

        for (const child of options.children) {
            this.root.push(new Math(child));
        }
    }
}
