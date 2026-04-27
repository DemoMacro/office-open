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
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

import type { IMargins } from "../floating";
import { TextWrappingSide } from "./text-wrapping";
import type { ITextWrapping } from "./text-wrapping";

interface IWrapThroughAttributes {
    readonly wrapText: (typeof TextWrappingSide)[keyof typeof TextWrappingSide];
    readonly distL?: number;
    readonly distR?: number;
}

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
const createWrapPolygon = (cx: number, cy: number): XmlComponent =>
    new BuilderElement({
        attributes: {
            edited: { key: "edited", value: "0" },
        },
        children: [
            new BuilderElement({
                attributes: {
                    x: { key: "x", value: 0 },
                    y: { key: "y", value: 0 },
                },
                name: "wp:start",
            }),
            new BuilderElement({
                attributes: {
                    x: { key: "x", value: 0 },
                    y: { key: "y", value: -cy },
                },
                name: "wp:lineTo",
            }),
            new BuilderElement({
                attributes: {
                    x: { key: "x", value: cx },
                    y: { key: "y", value: -cy },
                },
                name: "wp:lineTo",
            }),
            new BuilderElement({
                attributes: {
                    x: { key: "x", value: cx },
                    y: { key: "y", value: 0 },
                },
                name: "wp:lineTo",
            }),
            new BuilderElement({
                attributes: {
                    x: { key: "x", value: 0 },
                    y: { key: "y", value: 0 },
                },
                name: "wp:lineTo",
            }),
        ],
        name: "wp:wrapPolygon",
    });

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
    textWrapping: ITextWrapping,
    margins: IMargins = {
        bottom: 0,
        left: 0,
        right: 0,
        top: 0,
    },
    extent: { x: number; y: number },
): XmlComponent =>
    new BuilderElement<IWrapThroughAttributes>({
        attributes: {
            distL: { key: "distL", value: margins.left },
            distR: { key: "distR", value: margins.right },
            wrapText: { key: "wrapText", value: textWrapping.side || TextWrappingSide.BOTH_SIDES },
        },
        children: [createWrapPolygon(extent.x, extent.y)],
        name: "wp:wrapThrough",
    });
