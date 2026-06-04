/**
 * Math Matrix Properties module for Office MathML.
 *
 * This module provides properties for the matrix element (CT_MPr).
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_mPr-1.html
 *
 * @module
 */
import { BuilderElement, onOffObj } from "@file/xml-components";
import type { BuilderChild, XmlComponent } from "@file/xml-components";

/**
 * Options for math matrix properties.
 *
 * @see {@link createMathMatrixProperties}
 */
/**
 * Options for a matrix column (CT_MC).
 */
export interface MathMatrixColumnOptions {
  /** Number of columns this entry spans (1–255) */
  readonly count?: number;
  /** Column justification */
  readonly mcJc?: "left" | "center" | "right" | "inside" | "outside";
}

export interface MathMatrixPropertiesOptions {
  /** Base justification (vertical alignment of cells) */
  readonly baseJc?: "top" | "bot" | "center";
  /** Hide place holders */
  readonly plcHide?: boolean;
  /** Row spacing rule (0=single, 1=1.5 lines, 2=double, 3=exactly, 4=multiple) */
  readonly rSpRule?: 0 | 1 | 2 | 3 | 4;
  /** Column gap rule (0=single, 1=1.5 lines, 2=double, 3=exactly, 4=multiple) */
  readonly cGpRule?: 0 | 1 | 2 | 3 | 4;
  /** Row spacing */
  readonly rSp?: number;
  /** Column spacing */
  readonly cSp?: number;
  /** Column group spacing */
  readonly cGp?: number;
  /** Matrix column definitions */
  readonly mcs?: readonly MathMatrixColumnOptions[];
}

/**
 * Creates math matrix properties element (m:mPr).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_MPr">
 *   <xsd:sequence>
 *     <xsd:element name="baseJc" type="CT_YAlign" minOccurs="0"/>
 *     <xsd:element name="plcHide" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="rSpRule" type="CT_SpacingRule" minOccurs="0"/>
 *     <xsd:element name="cGpRule" type="CT_SpacingRule" minOccurs="0"/>
 *     <xsd:element name="rSp" type="CT_UnSignedInteger" minOccurs="0"/>
 *     <xsd:element name="cSp" type="CT_UnSignedInteger" minOccurs="0"/>
 *     <xsd:element name="cGp" type="CT_UnSignedInteger" minOccurs="0"/>
 *     <xsd:element name="mcs" type="CT_MCS" minOccurs="0"/>
 *     <xsd:element name="ctrlPr" type="CT_CtrlPr" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export const createMathMatrixProperties = (options: MathMatrixPropertiesOptions): XmlComponent => {
  const children: BuilderChild[] = [];

  if (options.baseJc !== undefined) {
    children.push(
      new BuilderElement({
        attributes: { val: { key: "m:val", value: options.baseJc } },
        name: "m:baseJc",
      }),
    );
  }
  if (options.plcHide !== undefined) {
    children.push(onOffObj("m:plcHide", options.plcHide));
  }
  if (options.rSpRule !== undefined) {
    children.push(
      new BuilderElement({
        attributes: { val: { key: "m:val", value: options.rSpRule } },
        name: "m:rSpRule",
      }),
    );
  }
  if (options.cGpRule !== undefined) {
    children.push(
      new BuilderElement({
        attributes: { val: { key: "m:val", value: options.cGpRule } },
        name: "m:cGpRule",
      }),
    );
  }
  if (options.rSp !== undefined) {
    children.push(
      new BuilderElement({
        attributes: { val: { key: "m:val", value: options.rSp } },
        name: "m:rSp",
      }),
    );
  }
  if (options.cSp !== undefined) {
    children.push(
      new BuilderElement({
        attributes: { val: { key: "m:val", value: options.cSp } },
        name: "m:cSp",
      }),
    );
  }
  if (options.cGp !== undefined) {
    children.push(
      new BuilderElement({
        attributes: { val: { key: "m:val", value: options.cGp } },
        name: "m:cGp",
      }),
    );
  }

  if (options.mcs !== undefined && options.mcs.length > 0) {
    children.push(
      new BuilderElement({
        children: options.mcs.map((mc) => {
          const mcPrChildren: BuilderChild[] = [];
          if (mc.count !== undefined) {
            mcPrChildren.push(
              new BuilderElement({
                attributes: { val: { key: "m:val", value: mc.count } },
                name: "m:count",
              }),
            );
          }
          if (mc.mcJc !== undefined) {
            mcPrChildren.push(
              new BuilderElement({
                attributes: { val: { key: "m:val", value: mc.mcJc } },
                name: "m:mcJc",
              }),
            );
          }
          const mcPr = new BuilderElement({ children: mcPrChildren, name: "m:mcPr" });
          return new BuilderElement({ children: [mcPr], name: "m:mc" });
        }),
        name: "m:mcs",
      }),
    );
  }

  return new BuilderElement({
    children,
    name: "m:mPr",
  });
};
