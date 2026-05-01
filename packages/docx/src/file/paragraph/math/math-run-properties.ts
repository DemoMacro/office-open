/**
 * Math Run Properties module for Office MathML.
 *
 * This module provides properties for math run elements (m:r).
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_rPr-1.html
 *
 * @module
 */
import { BuilderElement, OnOffElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

/**
 * Script type for math run properties.
 *
 * - "roman" - Roman script (default)
 * - "script" - Script (calligraphic)
 * - "fraktur" - Fraktur (Gothic)
 * - "double-struck" - Double-struck (blackboard bold)
 * - "sans-serif" - Sans-serif
 * - "monospace" - Monospace
 */
export type MathScriptType =
    | "roman"
    | "script"
    | "fraktur"
    | "double-struck"
    | "sans-serif"
    | "monospace";

/**
 * Style type for math run properties.
 *
 * - "p" - Plain (normal)
 * - "b" - Bold
 * - "i" - Italic
 * - "bi" - Bold italic
 */
export type MathStyleType = "p" | "b" | "i" | "bi";

/**
 * Options for math run properties.
 *
 * @see {@link createMathRunProperties}
 */
export interface MathRunPropertiesOptions {
    /** Whether the text is displayed in literal form */
    readonly lit?: boolean;
    /** Use normal text rendering instead of math italic */
    readonly normal?: boolean;
    /** Script type for the text */
    readonly script?: MathScriptType;
    /** Style type for the text */
    readonly style?: MathStyleType;
    /** Alignment point for manual break (1-255) */
    readonly breakAlignment?: number;
    /** Whether the text is aligned */
    readonly align?: boolean;
}

/**
 * Creates properties for a math run element (m:rPr).
 *
 * This element specifies run properties for math text,
 * including script style, literal rendering, and alignment.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_rPr-1.html
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_RPR">
 *   <xsd:sequence>
 *     <xsd:element name="lit" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:choice>
 *       <xsd:element name="nor" type="CT_OnOff" minOccurs="0"/>
 *       <xsd:sequence>
 *         <xsd:group ref="EG_ScriptStyle"/>  <!-- scr + sty -->
 *       </xsd:sequence>
 *     </xsd:choice>
 *     <xsd:element name="brk" type="CT_ManualBreak" minOccurs="0"/>
 *     <xsd:element name="aln" type="CT_OnOff" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export const createMathRunProperties = (options: MathRunPropertiesOptions): XmlComponent => {
    const children: XmlComponent[] = [];

    if (options.lit !== undefined) {
        children.push(new OnOffElement("m:lit", options.lit));
    }

    if (options.normal !== undefined) {
        children.push(new OnOffElement("m:nor", options.normal));
    } else if (options.script !== undefined || options.style !== undefined) {
        if (options.script !== undefined) {
            children.push(
                new BuilderElement({
                    attributes: { val: { key: "val", value: options.script } },
                    name: "m:scr",
                }),
            );
        }
        if (options.style !== undefined) {
            children.push(
                new BuilderElement({
                    attributes: { val: { key: "val", value: options.style } },
                    name: "m:sty",
                }),
            );
        }
    }

    if (options.breakAlignment !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { alnAt: { key: "m:alnAt", value: options.breakAlignment.toString() } },
                name: "m:brk",
            }),
        );
    }

    if (options.align !== undefined) {
        children.push(new OnOffElement("m:aln", options.align));
    }

    return new BuilderElement({
        children,
        name: "m:rPr",
    });
};
