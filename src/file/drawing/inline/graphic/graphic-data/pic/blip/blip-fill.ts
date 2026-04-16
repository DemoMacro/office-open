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
import { XmlComponent } from "@file/xml-components";

import { createBlip } from "./blip";
import { createSourceRectangle } from "./source-rectangle";
import { Stretch } from "./stretch";

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
export class BlipFill extends XmlComponent {
    public constructor(mediaData: IMediaData) {
        super("pic:blipFill");

        this.root.push(createBlip(mediaData));
        this.root.push(createSourceRectangle(mediaData.srcRect));
        this.root.push(new Stretch());
    }
}
