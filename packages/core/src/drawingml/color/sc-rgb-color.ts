/**
 * ScRGB color element for DrawingML shapes.
 *
 * This module provides scRGB color support using percentage-based RGB values.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_ScRgbColor
 *
 * @module
 */
import { element } from "@office-open/xml";

import type { ColorTransformOptions } from "./color-transform";
import { createColorTransforms } from "./color-transform";

/**
 * Options for scRGB color.
 */
export interface ScRgbColorOptions {
  /** Red percentage (e.g., "50%" or "100%") */
  r: string;
  /** Green percentage (e.g., "50%" or "100%") */
  g: string;
  /** Blue percentage (e.g., "50%" or "100%") */
  b: string;
  /** Optional color transforms */
  transforms?: ColorTransformOptions;
}

/**
 * Creates an scRGB color element as an XML string.
 *
 * Specifies a color using percentage-based RGB values.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_ScRgbColor">
 *   <xsd:sequence>
 *     <xsd:group ref="EG_ColorTransform" minOccurs="0" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="r" type="ST_Percentage" use="required"/>
 *   <xsd:attribute name="g" type="ST_Percentage" use="required"/>
 *   <xsd:attribute name="b" type="ST_Percentage" use="required"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * const redColor = createScRgbColor({ r: "100%", g: "0%", b: "0%" });
 * // With alpha transform
 * const semiRed = createScRgbColor({ r: "100%", g: "0%", b: "0%", transforms: { alpha: 50000 } });
 * ```
 */
export const createScRgbColor = (options: ScRgbColorOptions): string => {
  const transforms = options.transforms ? createColorTransforms(options.transforms) : [];
  return element("a:scrgbClr", { r: options.r, g: options.g, b: options.b }, transforms);
};
