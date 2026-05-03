import type { IExtendedMediaData } from "@file/media";
import { XmlComponent } from "@file/xml-components";
import type { FillOptions } from "@office-open/core/drawingml";

import { Anchor } from "./anchor";
import type { DocPropertiesOptions } from "./doc-properties/doc-properties";
import type { IFloating } from "./floating";
import { createInline } from "./inline";
import type { BlipEffectsOptions } from "./inline/graphic/graphic-data/pic/blip/blip-effects";
import type { TileOptions } from "./inline/graphic/graphic-data/pic/blip/tile";
import type { EffectListOptions } from "./inline/graphic/graphic-data/pic/shape-properties/effects/effect-list";
import type { OutlineOptions } from "./inline/graphic/graphic-data/pic/shape-properties/outline/outline";

/**
 * Distance options for drawing elements.
 *
 * Specifies the margins around a drawing element.
 */
export interface IDistance {
    readonly distT?: number;
    readonly distB?: number;
    readonly distL?: number;
    readonly distR?: number;
}

/**
 * Options for configuring a drawing element.
 *
 * @see {@link Drawing}
 */
export interface IDrawingOptions {
    readonly floating?: IFloating;
    readonly docProperties?: DocPropertiesOptions;
    readonly outline?: OutlineOptions;
    readonly fill?: FillOptions;
    readonly effects?: EffectListOptions;
    readonly blipEffects?: BlipEffectsOptions;
    readonly tile?: TileOptions;
}

/**
 * Represents a drawing element in a WordprocessingML document.
 *
 * Drawings can be either inline (positioned as part of the text flow) or
 * anchored (positioned relative to the page, column, or paragraph).
 *
 * Reference: http://officeopenxml.com/drwOverview.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Drawing">
 *   <xsd:choice minOccurs="1" maxOccurs="unbounded">
 *     <xsd:element ref="wp:anchor" minOccurs="0"/>
 *     <xsd:element ref="wp:inline" minOccurs="0"/>
 *   </xsd:choice>
 * </xsd:complexType>
 * ```
 */
export class Drawing extends XmlComponent {
    public constructor(imageData: IExtendedMediaData, drawingOptions: IDrawingOptions = {}) {
        super("w:drawing");

        if (!drawingOptions.floating) {
            this.root.push(
                createInline({
                    blipEffects: drawingOptions.blipEffects,
                    docProperties: drawingOptions.docProperties,
                    effects: drawingOptions.effects,
                    mediaData: imageData,
                    outline: drawingOptions.outline,
                    fill: drawingOptions.fill,
                    tile: drawingOptions.tile,
                    transform: imageData.transformation,
                }),
            );
        } else {
            this.root.push(
                new Anchor({
                    drawingOptions,
                    mediaData: imageData,
                    transform: imageData.transformation,
                }),
            );
        }
    }
}
