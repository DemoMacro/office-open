/**
 * Wrap Tight module for DrawingML text wrapping.
 *
 * This module provides tight text wrapping for floating drawings
 * where text wraps closely around the image shape.
 *
 * Reference: http://officeopenxml.com/drwPicFloating-textWrap.php
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

import type { IMargins } from "../floating";

interface IWrapTightAttributes {
    readonly distT?: number;
    readonly distB?: number;
}

/**
 * Creates tight text wrapping for a floating drawing.
 *
 * WrapTight causes text to wrap closely around the contours
 * of the drawing rather than its rectangular bounding box.
 *
 * Reference: http://officeopenxml.com/drwPicFloating-textWrap.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_WrapTight">
 *   <xsd:sequence>
 *     <xsd:element name="wrapPolygon" type="CT_WrapPath"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="wrapText" type="ST_WrapText" use="required"/>
 *   <xsd:attribute name="distL" type="ST_WrapDistance"/>
 *   <xsd:attribute name="distR" type="ST_WrapDistance"/>
 * </xsd:complexType>
 * ```
 */
export const createWrapTight = (
    margins: IMargins = {
        bottom: 0,
        top: 0,
    },
): XmlComponent =>
    new BuilderElement<IWrapTightAttributes>({
        attributes: {
            distB: { key: "distB", value: margins.bottom },
            distT: { key: "distT", value: margins.top },
        },
        name: "wp:wrapTight",
    });
