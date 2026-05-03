/**
 * Blip (Binary Large Image or Picture) module for DrawingML.
 *
 * This module provides the blip element that references the actual
 * image data within a picture.
 *
 * Reference: http://officeopenxml.com/drwPic.php
 *
 * @module
 */
import { BuilderElement } from "../../xml-components";
import type { XmlComponent } from "../../xml-components";
import { createBlipEffects } from "./blip-effects";
import type { BlipEffectsOptions } from "./blip-effects";
import { createExtentionList } from "./blip-extentions";

/**
 * Options for creating a blip element.
 */
export interface BlipOptions {
    /** Relationship ID of the embedded image (the value after "rId{...}") */
    readonly referenceId: string;
    /** Image type for SVG detection */
    readonly type?: "svg" | string;
    /** For SVG images, the fallback image file name */
    readonly fallbackFileName?: string;
}

/**
 * Creates a blip element for an image.
 *
 * A blip references the actual image data stored in the document package
 * through a relationship ID. For SVG images, it includes extensions that
 * reference the SVG data.
 *
 * Reference: http://officeopenxml.com/drwPic.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Blip">
 *   <xsd:sequence>
 *     <xsd:choice minOccurs="0" maxOccurs="unbounded">
 *       <xsd:element name="alphaBiLevel" type="CT_AlphaBiLevelEffect"/>
 *       <xsd:element name="alphaCeiling" type="CT_AlphaCeilingEffect"/>
 *       <xsd:element name="alphaFloor" type="CT_AlphaFloorEffect"/>
 *       <xsd:element name="alphaInv" type="CT_AlphaInverseEffect"/>
 *       <xsd:element name="alphaMod" type="CT_AlphaModulateEffect"/>
 *       <xsd:element name="alphaModFix" type="CT_AlphaModulateFixedEffect"/>
 *       <xsd:element name="alphaRepl" type="CT_AlphaReplaceEffect"/>
 *       <xsd:element name="biLevel" type="CT_BiLevelEffect"/>
 *       <xsd:element name="blur" type="CT_BlurEffect"/>
 *       <xsd:element name="clrChange" type="CT_ColorChangeEffect"/>
 *       <xsd:element name="clrRepl" type="CT_ColorReplaceEffect"/>
 *       <xsd:element name="duotone" type="CT_DuotoneEffect"/>
 *       <xsd:element name="fillOverlay" type="CT_FillOverlayEffect"/>
 *       <xsd:element name="grayscl" type="CT_GrayscaleEffect"/>
 *       <xsd:element name="hsl" type="CT_HSLEffect"/>
 *       <xsd:element name="lum" type="CT_LuminanceEffect"/>
 *       <xsd:element name="tint" type="CT_TintEffect"/>
 *     </xsd:choice>
 *     <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute ref="r:embed"/>
 *   <xsd:attribute ref="r:link"/>
 *   <xsd:attribute name="cstate" type="ST_BlipCompression"/>
 * </xsd:complexType>
 * ```
 *
 * @param options - Blip options including referenceId and type
 * @param blipEffects - Optional blip effects (brightness, contrast, etc.)
 * @returns An XML component representing the blip element
 */
export const createBlip = (
    options: BlipOptions,
    blipEffects?: BlipEffectsOptions,
): XmlComponent => {
    const children: XmlComponent[] = [];

    if (blipEffects) {
        children.push(...createBlipEffects(blipEffects));
    }

    if (options.type === "svg" && options.fallbackFileName) {
        children.push(createExtentionList(options.referenceId));
    }

    return new BuilderElement<{
        readonly embed: string;
        readonly cstate: string;
    }>({
        attributes: {
            cstate: {
                key: "cstate",
                value: "none",
            },
            embed: {
                key: "r:embed",
                value: `rId{${options.type === "svg" && options.fallbackFileName ? options.fallbackFileName : options.referenceId}}`,
            },
        },
        children,
        name: "a:blip",
    });
};
