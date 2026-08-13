/**
 * Tile fill module for blip fills.
 *
 * This module defines how images are tiled (repeated) to fill shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_TileInfoProperties
 *
 * @module
 */
import { element } from "@office-open/xml";

import { xsdRectAlignment } from "../../util/mappings";

/**
 * Tile flip mode for tiling images.
 *
 * Specifies whether the image is flipped along the x-axis, y-axis,
 * both axes, or not at all when tiling.
 */
export type TileFlipMode = "none" | "x" | "y" | "xy";

/**
 * Tile alignment within the shape.
 *
 * Specifies the anchor position of the first tile relative to the shape.
 */
export type TileAlignment =
  | "topLeft"
  | "top"
  | "topRight"
  | "left"
  | "center"
  | "right"
  | "bottomLeft"
  | "bottom"
  | "bottomRight";

/**
 * Options for tile fill mode.
 *
 * Configures how an image is tiled (repeated) to fill a shape.
 */
export interface TileOptions {
  /** Horizontal offset for the tile origin (in EMUs) */
  tx?: number;
  /** Vertical offset for the tile origin (in EMUs) */
  ty?: number;
  /** Horizontal scale as integer percent (100 = 100%) */
  sx?: number;
  /** Vertical scale as integer percent (100 = 100%) */
  sy?: number;
  /** Flip mode for alternating tiles */
  flip?: TileFlipMode;
  /** Alignment of the first tile within the shape */
  align?: TileAlignment;
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
 * createTileInfo({ flip: "xy", align: "center" });
 * ```
 */
export const createTileInfo = (options?: TileOptions): string => {
  if (!options) {
    return `<a:tile/>`;
  }

  const attrs: Record<string, string | number | undefined> = {};
  if (options.tx !== undefined) attrs.tx = options.tx;
  if (options.ty !== undefined) attrs.ty = options.ty;
  if (options.sx !== undefined) attrs.sx = options.sx * 1000;
  if (options.sy !== undefined) attrs.sy = options.sy * 1000;
  if (options.flip !== undefined) attrs.flip = options.flip;
  if (options.align !== undefined) attrs.algn = xsdRectAlignment.to(options.align);

  return element("a:tile", attrs);
};
