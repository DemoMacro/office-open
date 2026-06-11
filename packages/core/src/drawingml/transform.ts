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
import type { UniversalMeasure } from "@office-open/core";
import { convertToEmu } from "@office-open/core";
import { element } from "@office-open/xml";

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
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
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
  childOffsetX?: number | UniversalMeasure;
  childOffsetY?: number | UniversalMeasure;
  childExtentWidth?: number | UniversalMeasure;
  childExtentHeight?: number | UniversalMeasure;
}

function buildXfrmAttrs(
  options: Transform2DOptions,
): Record<string, string | number | boolean> | undefined {
  const attrs: Record<string, string | number | boolean> = {};
  if (options.flipHorizontal !== undefined) attrs.flipH = options.flipHorizontal;
  if (options.flipVertical !== undefined) attrs.flipV = options.flipVertical;
  if (options.rotation !== undefined) attrs.rot = options.rotation;
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
): string => {
  const children: string[] = [];

  if (options.x !== undefined || options.y !== undefined) {
    const x = options.x !== undefined ? convertToEmu(options.x) : 0;
    const y = options.y !== undefined ? convertToEmu(options.y) : 0;
    children.push(`<a:off x="${x}" y="${y}"/>`);
  }

  if (options.width !== undefined || options.height !== undefined) {
    const cx = options.width !== undefined ? convertToEmu(options.width) : 0;
    const cy = options.height !== undefined ? convertToEmu(options.height) : 0;
    children.push(`<a:ext cx="${cx}" cy="${cy}"/>`);
  }

  return element(elementName, buildXfrmAttrs(options), children);
};

/**
 * Creates a group transform element (a:xfrm with chOff/chExt children).
 */
export const createGroupTransform2D = (
  options: GroupTransform2DOptions,
  elementName: string = "a:xfrm",
): string => {
  const children: string[] = [];

  if (options.x !== undefined || options.y !== undefined) {
    const x = options.x !== undefined ? convertToEmu(options.x) : 0;
    const y = options.y !== undefined ? convertToEmu(options.y) : 0;
    children.push(`<a:off x="${x}" y="${y}"/>`);
  }

  if (options.width !== undefined || options.height !== undefined) {
    const cx = options.width !== undefined ? convertToEmu(options.width) : 0;
    const cy = options.height !== undefined ? convertToEmu(options.height) : 0;
    children.push(`<a:ext cx="${cx}" cy="${cy}"/>`);
  }

  const chOffX = options.childOffsetX !== undefined ? convertToEmu(options.childOffsetX) : 0;
  const chOffY = options.childOffsetY !== undefined ? convertToEmu(options.childOffsetY) : 0;
  children.push(`<a:chOff x="${chOffX}" y="${chOffY}"/>`);

  const chExtCx =
    options.childExtentWidth !== undefined ? convertToEmu(options.childExtentWidth) : 0;
  const chExtCy =
    options.childExtentHeight !== undefined ? convertToEmu(options.childExtentHeight) : 0;
  children.push(`<a:chExt cx="${chExtCx}" cy="${chExtCy}"/>`);

  return element(elementName, buildXfrmAttrs(options), children);
};
