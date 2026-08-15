/**
 * HSL color element for DrawingML.
 *
 * This module provides HSL (Hue, Saturation, Luminance) color support.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_HslColor
 *
 * @module
 */
import { element } from "@office-open/xml";

import type { ColorTransformOptions } from "./color-transform";
import { createColorTransforms } from "./color-transform";

/**
 * Options for HSL color.
 */
export interface HslColorOptions {
  /** Hue angle in degrees (0-360). */
  hue: number;
  /** Saturation as integer percent (0-100) */
  saturation: number;
  /** Luminance as integer percent (0-100) */
  luminance: number;
  /** Optional color transforms */
  transforms?: ColorTransformOptions;
}

/**
 * Creates an HSL color element as an XML string.
 *
 * Specifies a color using Hue, Saturation, and Luminance values.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_HslColor">
 *   <xsd:sequence>
 *     <xsd:group ref="EG_ColorTransform" minOccurs="0" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="hue" type="ST_PositiveFixedAngle" use="required"/>
 *   <xsd:attribute name="sat" type="ST_Percentage" use="required"/>
 *   <xsd:attribute name="lum" type="ST_Percentage" use="required"/>
 * </xsd:complexType>
 * ```
 */
export const createHslColor = (options: HslColorOptions): string => {
  const transforms = options.transforms ? createColorTransforms(options.transforms) : [];
  return element(
    "a:hslClr",
    {
      hue: Math.round(options.hue * 60000),
      sat: Math.round(options.saturation * 1000),
      lum: Math.round(options.luminance * 1000),
    },
    transforms,
  );
};
