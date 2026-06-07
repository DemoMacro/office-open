/**
 * Math Border Box Properties module for Office MathML.
 *
 * This module provides properties for the math border box element (CT_BorderBoxPr).
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_borderBoxPr-1.html
 *
 * @module
 */
import { BuilderElement, onOffObj } from "@file/xml-components";
import type { BuilderChild, XmlComponent } from "@file/xml-components";

/**
 * Options for math border box properties.
 *
 * @see {@link createMathBorderBoxProperties}
 */
export interface MathBorderBoxPropertiesOptions {
  /** Hide top border */
  hideTop?: boolean;
  /** Hide bottom border */
  hideBottom?: boolean;
  /** Hide left border */
  hideLeft?: boolean;
  /** Hide right border */
  hideRight?: boolean;
  /** Horizontal strike-through */
  strikeHorizontal?: boolean;
  /** Vertical strike-through */
  strikeVertical?: boolean;
  /** Bottom-left to top-right diagonal strike */
  strikeDiagonalUp?: boolean;
  /** Top-left to bottom-right diagonal strike */
  strikeDiagonalDown?: boolean;
}

/**
 * Creates math border box properties element (m:borderBoxPr).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_BorderBoxPr">
 *   <xsd:sequence>
 *     <xsd:element name="hideTop" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="hideBot" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="hideLeft" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="hideRight" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="strikeH" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="strikeV" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="strikeBLTR" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="strikeTLBR" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="ctrlPr" type="CT_CtrlPr" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export const createMathBorderBoxProperties = (
  options: MathBorderBoxPropertiesOptions,
): XmlComponent => {
  const children: BuilderChild[] = [];

  if (options.hideTop !== undefined) {
    children.push(onOffObj("m:hideTop", options.hideTop));
  }
  if (options.hideBottom !== undefined) {
    children.push(onOffObj("m:hideBot", options.hideBottom));
  }
  if (options.hideLeft !== undefined) {
    children.push(onOffObj("m:hideLeft", options.hideLeft));
  }
  if (options.hideRight !== undefined) {
    children.push(onOffObj("m:hideRight", options.hideRight));
  }
  if (options.strikeHorizontal !== undefined) {
    children.push(onOffObj("m:strikeH", options.strikeHorizontal));
  }
  if (options.strikeVertical !== undefined) {
    children.push(onOffObj("m:strikeV", options.strikeVertical));
  }
  if (options.strikeDiagonalUp !== undefined) {
    children.push(onOffObj("m:strikeBLTR", options.strikeDiagonalUp));
  }
  if (options.strikeDiagonalDown !== undefined) {
    children.push(onOffObj("m:strikeTLBR", options.strikeDiagonalDown));
  }

  return new BuilderElement({
    children,
    name: "m:borderBoxPr",
  });
};
