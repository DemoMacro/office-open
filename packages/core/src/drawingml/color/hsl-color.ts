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
  /** Hue angle in 60,000ths of a degree (0-21600000) */
  hue: number;
  /** Saturation in 1/1000th of a percent (0-100000) */
  saturation: number;
  /** Luminance in 1/1000th of a percent (0-100000) */
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
    { hue: options.hue, sat: options.saturation, lum: options.luminance },
    transforms,
  );
};
