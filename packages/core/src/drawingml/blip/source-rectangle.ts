/**
 * Source rectangle module for blip fills.
 *
 * This module defines the portion of an image to use when filling a shape.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_RelativeRect
 *
 * @module
 */
import { element } from "@office-open/xml";

/**
 * Options for source rectangle cropping.
 *
 * Each value is a percentage (0-100000) of the image dimension to crop.
 * The values represent the inset from each edge.
 */
export interface SourceRectangleOptions {
  /** Left inset percentage (0-100000) */
  left?: number;
  /** Top inset percentage (0-100000) */
  top?: number;
  /** Right inset percentage (0-100000) */
  right?: number;
  /** Bottom inset percentage (0-100000) */
  bottom?: number;
}

/**
 * Creates a source rectangle element for blip fill cropping.
 *
 * This element specifies a portion of the blip (image) to use as the fill.
 * When no options are provided, the entire blip is used.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_RelativeRect
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_RelativeRect">
 *   <xsd:attribute name="l" type="ST_Percentage" use="optional" default="0"/>
 *   <xsd:attribute name="t" type="ST_Percentage" use="optional" default="0"/>
 *   <xsd:attribute name="r" type="ST_Percentage" use="optional" default="0"/>
 *   <xsd:attribute name="b" type="ST_Percentage" use="optional" default="0"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Crop 10% from left and right
 * createSourceRectangle({ left: 10000, right: 10000 });
 * ```
 */
export const createSourceRectangle = (options?: SourceRectangleOptions): string => {
  if (!options) {
    return element("a:srcRect");
  }

  const attrs: Record<string, string | number | undefined> = {};
  if (options.left !== undefined) attrs.l = options.left;
  if (options.top !== undefined) attrs.t = options.top;
  if (options.right !== undefined) attrs.r = options.right;
  if (options.bottom !== undefined) attrs.b = options.bottom;

  return element("a:srcRect", attrs);
};
