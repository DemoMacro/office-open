/**
 * Math Base module for Office MathML.
 *
 * This module provides the base element (m:e) that contains
 * the primary content of math structures.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_e-1.html
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { BuilderChild, XmlComponent } from "@file/xml-components";

import type { MathComponent } from "../math-component";

/**
 * Options for math argument properties (CT_OMathArgPr).
 */
export interface MathArgPropertiesOptions {
  /** Argument size (-2 to 2) */
  readonly argSz?: number;
}

/**
 * Options for creating a math base element.
 */
export interface MathBaseOptions {
  /** The content of the base */
  readonly children: readonly MathComponent[];
  /** Argument properties */
  readonly argPr?: MathArgPropertiesOptions;
}

/**
 * Creates a math arg properties element (m:argPr).
 */
const createMathArgProperties = (options: MathArgPropertiesOptions): XmlComponent => {
  const children: BuilderChild[] = [];
  if (options.argSz !== undefined) {
    children.push(
      new BuilderElement({
        attributes: { val: { key: "m:val", value: options.argSz } },
        name: "m:argSz",
      }),
    );
  }
  return new BuilderElement({ children, name: "m:argPr" });
};

/**
 * Creates a math base element.
 *
 * The math base (m:e) is used within math structures like fractions,
 * radicals, and n-ary operators to contain the primary expression.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_e-1.html
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_OMathArg">
 *   <xsd:sequence>
 *     <xsd:element name="argPr" type="CT_OMathArgPr" minOccurs="0"/>
 *     <xsd:group ref="EG_OMathMathElements" minOccurs="0" maxOccurs="unbounded"/>
 *     <xsd:element name="ctrlPr" type="CT_CtrlPr" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export const createMathBase = ({ children, argPr }: MathBaseOptions): XmlComponent => {
  const allChildren: BuilderChild[] = [];
  if (argPr) {
    allChildren.push(createMathArgProperties(argPr));
  }
  allChildren.push(...children);
  return new BuilderElement({
    children: allChildren,
    name: "m:e",
  });
};
