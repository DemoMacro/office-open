/**
 * Tile fill module for blip fills.
 *
 * This module defines how images are tiled (repeated) to fill shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_TileInfoProperties
 *
 * @module
 */
import { BuilderElement } from "../../xml-components";

/**
 * Tile flip mode for tiling images.
 *
 * Specifies whether the image is flipped along the x-axis, y-axis,
 * both axes, or not at all when tiling.
 */
export const TileFlipMode = {
    /** No flipping */
    NONE: "none",
    /** Flip along x-axis */
    X: "x",
    /** Flip along y-axis */
    Y: "y",
    /** Flip along both axes */
    XY: "xy",
} as const;

/**
 * Tile alignment within the shape.
 *
 * Specifies the anchor position of the first tile relative to the shape.
 */
export const TileAlignment = {
    /** Top-left corner */
    TOP_LEFT: "tl",
    /** Top center */
    TOP: "t",
    /** Top-right corner */
    TOP_RIGHT: "tr",
    /** Middle-left */
    LEFT: "l",
    /** Center */
    CENTER: "ctr",
    /** Middle-right */
    RIGHT: "r",
    /** Bottom-left corner */
    BOTTOM_LEFT: "bl",
    /** Bottom center */
    BOTTOM: "b",
    /** Bottom-right corner */
    BOTTOM_RIGHT: "br",
} as const;

/**
 * Options for tile fill mode.
 *
 * Configures how an image is tiled (repeated) to fill a shape.
 */
export interface TileOptions {
    /** Horizontal offset for the tile origin (in EMUs) */
    readonly tx?: number;
    /** Vertical offset for the tile origin (in EMUs) */
    readonly ty?: number;
    /** Horizontal scale factor as percentage (e.g., 50 = 50%) */
    readonly sx?: number;
    /** Vertical scale factor as percentage (e.g., 50 = 50%) */
    readonly sy?: number;
    /** Flip mode for alternating tiles */
    readonly flip?: keyof typeof TileFlipMode;
    /** Alignment of the first tile within the shape */
    readonly align?: keyof typeof TileAlignment;
}

/**
 * Creates a tile fill mode element for blip fills.
 *
 * When a blip fill uses tile mode, the image is repeated (tiled) to fill
 * the shape instead of being stretched. This element controls the tiling
 * parameters such as offset, scale, flip, and alignment.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_TileInfoProperties
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_TileInfoProperties">
 *   <xsd:attribute name="tx" type="ST_Coordinate" use="optional"/>
 *   <xsd:attribute name="ty" type="ST_Coordinate" use="optional"/>
 *   <xsd:attribute name="sx" type="ST_Percentage" use="optional"/>
 *   <xsd:attribute name="sy" type="ST_Percentage" use="optional"/>
 *   <xsd:attribute name="flip" type="ST_TileFlipMode" default="none"/>
 *   <xsd:attribute name="algn" type="ST_RectAlignment" use="optional"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Tile with 50% scale
 * createTileInfo({ sx: 50, sy: 50 });
 * // Tile with flip and alignment
 * createTileInfo({ flip: "XY", align: "CENTER" });
 * ```
 */
export const createTileInfo = (options?: TileOptions) => {
    if (!options) {
        return new BuilderElement({ name: "a:tile" });
    }

    const attributes: Record<string, { readonly key: string; readonly value: string | number }> =
        {};

    if (options.tx !== undefined) {
        attributes.tx = { key: "tx", value: options.tx };
    }
    if (options.ty !== undefined) {
        attributes.ty = { key: "ty", value: options.ty };
    }
    if (options.sx !== undefined) {
        attributes.sx = { key: "sx", value: options.sx };
    }
    if (options.sy !== undefined) {
        attributes.sy = { key: "sy", value: options.sy };
    }
    if (options.flip !== undefined) {
        attributes.flip = { key: "flip", value: TileFlipMode[options.flip] };
    }
    if (options.align !== undefined) {
        attributes.algn = { key: "algn", value: TileAlignment[options.align] };
    }

    return new BuilderElement({
        attributes: Object.keys(attributes).length > 0 ? (attributes as never) : undefined,
        name: "a:tile",
    });
};
