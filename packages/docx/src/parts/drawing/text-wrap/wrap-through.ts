/**
 * Wrap Through module for DrawingML text wrapping.
 *
 * This module provides "through" text wrapping for floating drawings
 * where text wraps through the image contours, filling any concave areas.
 *
 * Reference: http://officeopenxml.com/drwPicFloating-textWrap.php
 *
 * @module
 */
import { element } from "@office-open/xml";

import type { Margins } from "../floating";
import { TextWrappingSide } from "./text-wrapping";
import type { TextWrapping } from "./text-wrapping";

/**
 * Creates a default rectangular wrap polygon matching the image extent.
 *
 * Reference: http://officeopenxml.com/drwPicFloating-textWrap.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_WrapPath">
 *   <xsd:sequence>
 *     <xsd:element name="start" type="a:CT_Point2D" minOccurs="1" maxOccurs="1"/>
 *     <xsd:element name="lineTo" type="a:CT_Point2D" minOccurs="2" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="edited" type="xsd:boolean" use="optional"/>
 * </xsd:complexType>
 * ```
 */
const createWrapPolygon = (cx: number, cy: number): string =>
  element("wp:wrapPolygon", { edited: "0" }, [
    `<wp:start x="0" y="0"/>`,
    `<wp:lineTo x="0" y="${-cy}"/>`,
    `<wp:lineTo x="${cx}" y="${-cy}"/>`,
    `<wp:lineTo x="${cx}" y="0"/>`,
    `<wp:lineTo x="0" y="0"/>`,
  ]);

/**
 * Creates "through" text wrapping for a floating drawing.
 *
 * WrapThrough is similar to WrapTight but allows text to wrap through
 * the concave portions of the drawing shape (e.g., the inside of the letter "O").
 * A default rectangular wrap polygon matching the image extent is generated.
 *
 * Reference: http://officeopenxml.com/drwPicFloating-textWrap.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_WrapThrough">
 *   <xsd:sequence>
 *     <xsd:element name="wrapPolygon" type="CT_WrapPath" minOccurs="1" maxOccurs="1"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="wrapText" type="ST_WrapText" use="required"/>
 *   <xsd:attribute name="distL" type="ST_WrapDistance"/>
 *   <xsd:attribute name="distR" type="ST_WrapDistance"/>
 * </xsd:complexType>
 * ```
 */
export const createWrapThrough = (
  textWrapping: TextWrapping,
  margins: Margins = {
    bottom: 0,
    left: 0,
    right: 0,
    top: 0,
  },
  extent: { x: number; y: number },
): string =>
  element(
    "wp:wrapThrough",
    {
      distL: margins.left,
      distR: margins.right,
      wrapText: textWrapping.side || TextWrappingSide.BOTH_SIDES,
    },
    [createWrapPolygon(extent.x, extent.y)],
  );
