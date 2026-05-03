/**
 * Picture module for DrawingML elements.
 *
 * This module provides the picture element that represents an image
 * within a DrawingML graphic.
 *
 * Reference: http://officeopenxml.com/drwPic.php
 *
 * @module
 */
// http://officeopenxml.com/drwPic.php
import type { HyperlinkOptions } from "@file/drawing/doc-properties/doc-properties";
import type { IMediaData, IMediaDataTransformation } from "@file/media";
import { XmlComponent } from "@file/xml-components";
import type { FillOptions, TileOptions } from "@office-open/core/drawingml";

import type { BlipOptions } from "./blip/blip";
import type { BlipEffectsOptions } from "./blip/blip-effects";
import { createBlipFill } from "./blip/blip-fill";
import { NonVisualPicProperties } from "./non-visual-pic-properties/non-visual-pic-properties";
import { PicAttributes } from "./pic-attributes";
import type { EffectListOptions } from "./shape-properties/effects/effect-list";
import type { OutlineOptions } from "./shape-properties/outline/outline";
import { ShapeProperties } from "./shape-properties/shape-properties";

/**
 * Represents a picture element in DrawingML.
 *
 * A picture contains non-visual properties, image data (blip fill),
 * and shape properties that define how the image is displayed.
 *
 * Reference: http://officeopenxml.com/drwPic.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Picture">
 *   <xsd:sequence>
 *     <xsd:element name="nvPicPr" type="CT_PictureNonVisual"/>
 *     <xsd:element name="blipFill" type="CT_BlipFillProperties"/>
 *     <xsd:element name="spPr" type="a:CT_ShapeProperties"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * const pic = new Pic({
 *   mediaData: imageData,
 *   transform: { emus: { x: 914400, y: 914400 } },
 *   outline: { color: "000000", width: 9525 }
 * });
 * ```
 */
export class Pic extends XmlComponent {
    public constructor({
        mediaData,
        transform,
        outline,
        fill,
        effects,
        blipEffects,
        tile,
        hyperlink,
    }: {
        readonly mediaData: IMediaData;
        readonly transform: IMediaDataTransformation;
        readonly outline?: OutlineOptions;
        readonly fill?: FillOptions;
        readonly effects?: EffectListOptions;
        readonly blipEffects?: BlipEffectsOptions;
        readonly tile?: TileOptions;
        readonly hyperlink?: HyperlinkOptions;
    }) {
        super("pic:pic");

        this.root.push(
            new PicAttributes({
                xmlns: "http://schemas.openxmlformats.org/drawingml/2006/picture",
            }),
        );

        this.root.push(new NonVisualPicProperties(hyperlink));
        const blipOptions: BlipOptions = {
            referenceId: mediaData.fileName,
            type: mediaData.type,
            fallbackFileName: "fallback" in mediaData ? mediaData.fallback.fileName : undefined,
        };
        this.root.push(
            createBlipFill(blipOptions, { blipEffects, tile, srcRect: mediaData.srcRect }),
        );
        this.root.push(new ShapeProperties({ element: "pic", effects, fill, outline, transform }));
    }
}
