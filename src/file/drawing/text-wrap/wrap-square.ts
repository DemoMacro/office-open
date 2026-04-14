/**
 * Wrap Square module for DrawingML text wrapping.
 *
 * This module provides square text wrapping for floating drawings
 * where text wraps around a rectangular bounding box.
 *
 * Reference: http://officeopenxml.com/drwPicFloating-textWrap.php
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

import type { IDistance } from "../drawing";
import type { IMargins } from "../floating";
import { TextWrappingSide } from "./text-wrapping";
import type { ITextWrapping } from "./text-wrapping";

type IWrapSquareAttributes = {
    readonly wrapText?: (typeof TextWrappingSide)[keyof typeof TextWrappingSide];
} & IDistance;

/**
 * Creates square text wrapping for a floating drawing.
 *
 * WrapSquare causes text to wrap around the rectangular bounding box
 * of the drawing on the specified side(s).
 *
 * Reference: http://officeopenxml.com/drwPicFloating-textWrap.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_WrapSquare">
 *   <xsd:sequence>
 *     <xsd:element name="effectExtent" type="CT_EffectExtent" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="wrapText" type="ST_WrapText" use="required"/>
 *   <xsd:attribute name="distT" type="ST_WrapDistance"/>
 *   <xsd:attribute name="distB" type="ST_WrapDistance"/>
 *   <xsd:attribute name="distL" type="ST_WrapDistance"/>
 *   <xsd:attribute name="distR" type="ST_WrapDistance"/>
 * </xsd:complexType>
 * ```
 */
export const createWrapSquare = (
    textWrapping: ITextWrapping,
    margins: IMargins = {
        bottom: 0,
        left: 0,
        right: 0,
        top: 0,
    },
): XmlComponent =>
    new BuilderElement<IWrapSquareAttributes>({
        attributes: {
            distB: { key: "distB", value: margins.bottom },
            distL: { key: "distL", value: margins.left },
            distR: { key: "distR", value: margins.right },
            distT: { key: "distT", value: margins.top },
            wrapText: { key: "wrapText", value: textWrapping.side || TextWrappingSide.BOTH_SIDES },
        },
        name: "wp:wrapSquare",
    });
