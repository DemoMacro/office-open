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
import { BuilderElement } from "../../xml-components";
import type { XmlComponent } from "../../xml-components";
import { createBlip } from "./blip";
import type { BlipOptions } from "./blip";
import type { BlipEffectsOptions } from "./blip-effects";
import { createSourceRectangle } from "./source-rectangle";
import type { SourceRectangleOptions } from "./source-rectangle";
import { Stretch } from "./stretch";
import { createTileInfo } from "./tile";
import type { TileOptions } from "./tile";

/**
 * Options for blip fill properties.
 */
export interface BlipFillOptions {
    /** DPI of the image */
    readonly dpi?: number;
    /** Whether the fill rotates with the shape */
    readonly rotWithShape?: boolean;
    /** Image adjustment effects (brightness, contrast, grayscale, etc.) */
    readonly blipEffects?: BlipEffectsOptions;
    /** Source rectangle for cropping */
    readonly srcRect?: SourceRectangleOptions;
    /** Tile fill mode (if omitted, defaults to stretch) */
    readonly tile?: TileOptions;
}

/**
 * Creates a blip fill element.
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
 * @param blipOptions - Blip options including referenceId and type
 * @param fillOptions - Optional blip fill options
 */
export const createBlipFill = (
    blipOptions: BlipOptions,
    fillOptions?: BlipFillOptions,
): XmlComponent => {
    const children: XmlComponent[] = [];

    children.push(createBlip(blipOptions, fillOptions?.blipEffects));
    children.push(createSourceRectangle(fillOptions?.srcRect));

    if (fillOptions?.tile) {
        children.push(createTileInfo(fillOptions.tile));
    } else {
        children.push(new Stretch());
    }

    const attributes: Record<string, { readonly key: string; readonly value: number }> = {};

    if (fillOptions?.dpi !== undefined) {
        attributes.dpi = { key: "dpi", value: fillOptions.dpi };
    }
    if (fillOptions?.rotWithShape !== undefined) {
        attributes.rotWithShape = {
            key: "rotWithShape",
            value: fillOptions.rotWithShape ? 1 : 0,
        };
    }

    return new BuilderElement({
        attributes: Object.keys(attributes).length > 0 ? (attributes as never) : undefined,
        children,
        name: "pic:blipFill",
    });
};
