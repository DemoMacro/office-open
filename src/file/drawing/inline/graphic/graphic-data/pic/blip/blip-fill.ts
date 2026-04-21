/**
 * Blip fill module for DrawingML pictures.
 *
 * This module defines how an image (blip) fills a picture shape,
 * including stretching and cropping options.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_BlipFillProperties
 *
 * @module
 */
import type { IMediaData } from "@file/media";
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

import { createBlip } from "./blip";
import { createSourceRectangle } from "./source-rectangle";
import { Stretch } from "./stretch";

/**
 * Options for blip fill properties.
 */
export interface BlipFillOptions {
    /** DPI of the image */
    readonly dpi?: number;
    /** Whether the fill rotates with the shape */
    readonly rotWithShape?: boolean;
}

/**
 * Represents a blip fill for pictures in DrawingML.
 *
 * This element specifies the type of fill used for a picture. It contains the blip (image)
 * reference, an optional source rectangle for cropping, and the fill mode (typically stretch).
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_BlipFillProperties
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_BlipFillProperties">
 *   <xsd:sequence>
 *     <xsd:element name="blip" type="CT_Blip" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="srcRect" type="CT_RelativeRect" minOccurs="0" maxOccurs="1"/>
 *     <xsd:group ref="EG_FillModeProperties" minOccurs="0" maxOccurs="1"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="dpi" type="xsd:unsignedInt" use="optional"/>
 *   <xsd:attribute name="rotWithShape" type="xsd:boolean" use="optional"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * const blipFill = new BlipFill(mediaData);
 * // If mediaData.srcRect is set, cropping is applied
 * ```
 */
export const createBlipFill = (mediaData: IMediaData, options?: BlipFillOptions): XmlComponent => {
    const children: XmlComponent[] = [];

    children.push(createBlip(mediaData));
    children.push(createSourceRectangle(mediaData.srcRect));
    children.push(new Stretch());

    const attributes: Record<string, { readonly key: string; readonly value: number }> = {};

    if (options?.dpi !== undefined) {
        attributes.dpi = { key: "dpi", value: options.dpi };
    }
    if (options?.rotWithShape !== undefined) {
        attributes.rotWithShape = {
            key: "rotWithShape",
            value: options.rotWithShape ? 1 : 0,
        };
    }

    return new BuilderElement({
        attributes: Object.keys(attributes).length > 0 ? (attributes as never) : undefined,
        children,
        name: "pic:blipFill",
    });
};
