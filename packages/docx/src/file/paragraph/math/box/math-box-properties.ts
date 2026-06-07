/**
 * Math Box Properties module for Office MathML.
 *
 * This module provides properties for the math box element (CT_BoxPr).
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_boxPr-1.html
 *
 * @module
 */
import { BuilderElement, onOffObj } from "@file/xml-components";
import type { BuilderChild, XmlComponent } from "@file/xml-components";

/**
 * Options for math box properties.
 *
 * @see {@link createMathBoxProperties}
 */
export interface MathBoxPropertiesOptions {
  /** Operator emulation */
  opEmu?: boolean;
  /** No break */
  noBreak?: boolean;
  /** Differential */
  diff?: boolean;
  /** Alignment */
  aln?: boolean;
}

/**
 * Creates math box properties element (m:boxPr).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_BoxPr">
 *   <xsd:sequence>
 *     <xsd:element name="opEmu" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="noBreak" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="diff" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="brk" type="CT_ManualBreak" minOccurs="0"/>
 *     <xsd:element name="aln" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="ctrlPr" type="CT_CtrlPr" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export const createMathBoxProperties = (options: MathBoxPropertiesOptions): XmlComponent => {
  const children: BuilderChild[] = [];

  if (options.opEmu !== undefined) {
    children.push(onOffObj("m:opEmu", options.opEmu));
  }
  if (options.noBreak !== undefined) {
    children.push(onOffObj("m:noBreak", options.noBreak));
  }
  if (options.diff !== undefined) {
    children.push(onOffObj("m:diff", options.diff));
  }
  if (options.aln !== undefined) {
    children.push(onOffObj("m:aln", options.aln));
  }

  return new BuilderElement({
    children,
    name: "m:boxPr",
  });
};
