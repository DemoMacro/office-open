/**
 * 2D transform for DrawingML shapes.
 *
 * This module provides factory functions for creating transform elements
 * (a:xfrm) used in shape properties, picture properties, and group shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_Transform2D / CT_GroupTransform2D
 *
 * @module
 */
import { BuilderElement } from "../xml-components";
import type { XmlComponent } from "../xml-components";

// <xsd:complexType name="CT_Transform2D">
//     <xsd:sequence>
//         <xsd:element name="off" type="CT_Point2D" minOccurs="0"/>
//         <xsd:element name="ext" type="CT_PositiveSize2D" minOccurs="0"/>
//     </xsd:sequence>
//     <xsd:attribute name="rot" type="ST_Angle" use="optional"/>
//     <xsd:attribute name="flipH" type="xsd:boolean" use="optional"/>
//     <xsd:attribute name="flipV" type="xsd:boolean" use="optional"/>
// </xsd:complexType>

export interface Transform2DOptions {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  flipHorizontal?: boolean;
  flipVertical?: boolean;
  rotation?: number;
}

// <xsd:complexType name="CT_GroupTransform2D">
//     <xsd:complexContent>
//         <xsd:extension base="CT_Transform2D">
//             <xsd:sequence>
//                 <xsd:element name="chOff" type="CT_Point2D" minOccurs="0"/>
//                 <xsd:element name="chExt" type="CT_PositiveSize2D" minOccurs="0"/>
//             </xsd:sequence>
//         </xsd:extension>
//     </xsd:complexContent>
// </xsd:complexType>

export interface GroupTransform2DOptions extends Transform2DOptions {
  childOffsetX?: number;
  childOffsetY?: number;
  childExtentWidth?: number;
  childExtentHeight?: number;
}

function buildXfrmAttrs(options: Transform2DOptions) {
  const attrs: Record<
    string,
    { readonly key: string; readonly value: string | number | boolean | undefined }
  > = {};
  if (options.flipHorizontal !== undefined)
    attrs.flipH = { key: "flipH", value: options.flipHorizontal };
  if (options.flipVertical !== undefined)
    attrs.flipV = { key: "flipV", value: options.flipVertical };
  if (options.rotation !== undefined) attrs.rot = { key: "rot", value: options.rotation };
  return Object.keys(attrs).length > 0 ? attrs : undefined;
}

/**
 * Creates a 2D transform element (a:xfrm).
 *
 * @param options - Transform options including position, size, rotation, and flip.
 * @param elementName - Element name, defaults to "a:xfrm".
 */
export const createTransform2D = (
  options: Transform2DOptions,
  elementName: string = "a:xfrm",
): XmlComponent => {
  const children: XmlComponent[] = [];

  if (options.x !== undefined || options.y !== undefined) {
    children.push(
      new BuilderElement({
        name: "a:off",
        attributes: {
          x: { key: "x", value: options.x ?? 0 },
          y: { key: "y", value: options.y ?? 0 },
        },
      }),
    );
  }

  if (options.width !== undefined || options.height !== undefined) {
    children.push(
      new BuilderElement({
        name: "a:ext",
        attributes: {
          cx: { key: "cx", value: options.width ?? 0 },
          cy: { key: "cy", value: options.height ?? 0 },
        },
      }),
    );
  }

  return new BuilderElement({
    name: elementName,
    attributes: buildXfrmAttrs(options) as never,
    children: children.length > 0 ? children : undefined,
  });
};

/**
 * Creates a group transform element (a:xfrm with chOff/chExt children).
 */
export const createGroupTransform2D = (
  options: GroupTransform2DOptions,
  elementName: string = "a:xfrm",
): XmlComponent => {
  const base = createTransform2D(options, elementName);

  base["root"].push(
    new BuilderElement({
      name: "a:chOff",
      attributes: {
        x: { key: "x", value: options.childOffsetX ?? 0 },
        y: { key: "y", value: options.childOffsetY ?? 0 },
      },
    }),
  );

  base["root"].push(
    new BuilderElement({
      name: "a:chExt",
      attributes: {
        cx: { key: "cx", value: options.childExtentWidth ?? 0 },
        cy: { key: "cy", value: options.childExtentHeight ?? 0 },
      },
    }),
  );

  return base;
};
