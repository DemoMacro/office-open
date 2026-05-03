/**
 * Math Properties module for Office MathML.
 *
 * This module provides global math document properties (m:mathPr).
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_mathPr-1.html
 *
 * @module
 */
import { BuilderElement, OnOffElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

/**
 * Binary operator break position.
 *
 * - "before" - Break before binary operator
 * - "after" - Break after binary operator
 * - "repeat" - Repeat binary operator on each line
 */
export type MathBreakBin = "before" | "after" | "repeat";

/**
 * Binary operator subtraction break position.
 *
 * - "--" - Minus on both lines
 * - "-+" - Minus on first, plus on second
 * - "+-" - Plus on first, minus on second
 */
export type MathBreakBinSub = "--" | "-+" | "+-";

/**
 * Math justification type for global math properties.
 *
 * - "left" - Left justify
 * - "right" - Right justify
 * - "center" - Center
 * - "centerGroup" - Center group
 */
export type MathPropertiesJustification = "left" | "right" | "center" | "centerGroup";

/**
 * Options for global math properties.
 *
 * @see {@link createMathProperties}
 */
export interface MathPropertiesOptions {
    /** Math font name */
    readonly mathFont?: string;
    /** Binary operator break position */
    readonly breakBin?: MathBreakBin;
    /** Binary operator subtraction break position */
    readonly breakBinSub?: MathBreakBinSub;
    /** Use small fractions */
    readonly smallFrac?: boolean;
    /** Display default */
    readonly displayDefault?: boolean;
    /** Left margin in twips */
    readonly leftMargin?: number;
    /** Right margin in twips */
    readonly rightMargin?: number;
    /** Default justification */
    readonly defaultJustification?: MathPropertiesJustification;
    /** Pre-spacing in twips */
    readonly preSpacing?: number;
    /** Post-spacing in twips */
    readonly postSpacing?: number;
    /** Inter-spacing in twips */
    readonly interSpacing?: number;
    /** Intra-spacing in twips */
    readonly intraSpacing?: number;
    /** Wrap indent in twips */
    readonly wrapIndent?: number;
    /** Wrap right */
    readonly wrapRight?: boolean;
    /** Integral limit location: "undOvr" (under/over) or "subSup" (subscript/superscript) */
    readonly integralLimitLocation?: string;
    /** N-ary limit location: "undOvr" (under/over) or "subSup" (subscript/superscript) */
    readonly naryLimitLocation?: string;
}

/**
 * Creates global math properties element (m:mathPr).
 *
 * This element specifies global math properties for the document,
 * including fonts, spacing, justification, and limit locations.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_mathPr-1.html
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_MathPr">
 *   <xsd:sequence>
 *     <xsd:element name="mathFont" type="CT_String" minOccurs="0"/>
 *     <xsd:element name="brkBin" type="CT_BreakBin" minOccurs="0"/>
 *     <xsd:element name="brkBinSub" type="CT_BreakBinSub" minOccurs="0"/>
 *     <xsd:element name="smallFrac" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="dispDef" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="lMargin" type="CT_TwipsMeasure" minOccurs="0"/>
 *     <xsd:element name="rMargin" type="CT_TwipsMeasure" minOccurs="0"/>
 *     <xsd:element name="defJc" type="CT_OMathJc" minOccurs="0"/>
 *     <xsd:element name="preSp" type="CT_TwipsMeasure" minOccurs="0"/>
 *     <xsd:element name="postSp" type="CT_TwipsMeasure" minOccurs="0"/>
 *     <xsd:element name="interSp" type="CT_TwipsMeasure" minOccurs="0"/>
 *     <xsd:element name="intraSp" type="CT_TwipsMeasure" minOccurs="0"/>
 *     <xsd:choice minOccurs="0">
 *       <xsd:element name="wrapIndent" type="CT_TwipsMeasure"/>
 *       <xsd:element name="wrapRight" type="CT_OnOff"/>
 *     </xsd:choice>
 *     <xsd:element name="intLim" type="CT_LimLoc" minOccurs="0"/>
 *     <xsd:element name="naryLim" type="CT_LimLoc" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export const createMathProperties = (options: MathPropertiesOptions): XmlComponent => {
    const children: XmlComponent[] = [];

    if (options.mathFont !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: options.mathFont } },
                name: "m:mathFont",
            }),
        );
    }
    if (options.breakBin !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: options.breakBin } },
                name: "m:brkBin",
            }),
        );
    }
    if (options.breakBinSub !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: options.breakBinSub } },
                name: "m:brkBinSub",
            }),
        );
    }
    if (options.smallFrac !== undefined) {
        children.push(new OnOffElement("m:smallFrac", options.smallFrac));
    }
    if (options.displayDefault !== undefined) {
        children.push(new OnOffElement("m:dispDef", options.displayDefault));
    }
    if (options.leftMargin !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: options.leftMargin.toString() } },
                name: "m:lMargin",
            }),
        );
    }
    if (options.rightMargin !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: options.rightMargin.toString() } },
                name: "m:rMargin",
            }),
        );
    }
    if (options.defaultJustification !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: options.defaultJustification } },
                name: "m:defJc",
            }),
        );
    }
    if (options.preSpacing !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: options.preSpacing.toString() } },
                name: "m:preSp",
            }),
        );
    }
    if (options.postSpacing !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: options.postSpacing.toString() } },
                name: "m:postSp",
            }),
        );
    }
    if (options.interSpacing !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: options.interSpacing.toString() } },
                name: "m:interSp",
            }),
        );
    }
    if (options.intraSpacing !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: options.intraSpacing.toString() } },
                name: "m:intraSp",
            }),
        );
    }
    if (options.wrapIndent !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: options.wrapIndent.toString() } },
                name: "m:wrapIndent",
            }),
        );
    }
    if (options.wrapRight !== undefined) {
        children.push(new OnOffElement("m:wrapRight", options.wrapRight));
    }
    if (options.integralLimitLocation !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: options.integralLimitLocation } },
                name: "m:intLim",
            }),
        );
    }
    if (options.naryLimitLocation !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { val: { key: "m:val", value: options.naryLimitLocation } },
                name: "m:naryLim",
            }),
        );
    }

    return new BuilderElement({
        children,
        name: "m:mathPr",
    });
};
