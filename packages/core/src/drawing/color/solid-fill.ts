/**
 * Solid fill element for DrawingML shapes.
 *
 * This module provides solid fill support for outlines and shapes,
 * supporting RGB, scheme, HSL, system, and preset colors.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_SolidColorFillProperties
 *
 * @module
 */
import { element } from "@office-open/xml";

import { emitColorChoice } from "./color-descriptors";
import type { HslColorOptions } from "./hsl-color";
import type { PresetColorOptions } from "./preset-color";
import type { RgbColorOptions } from "./rgb-color";
import type { ScRgbColorOptions } from "./sc-rgb-color";
import type { SchemeColorOptions } from "./scheme-color";
import type { SystemColorOptions } from "./system-color";

/**
 * Union type for all color options supported by solid fill.
 *
 * Extends the original pattern with additional color types:
 * RGB, scheme, HSL, system, and preset colors.
 */
export type SolidFillOptions =
  | ScRgbColorOptions
  | RgbColorOptions
  | SchemeColorOptions
  | HslColorOptions
  | SystemColorOptions
  | PresetColorOptions;

/**
 * Creates the color child element for a solid fill — thin delegate to the
 * single EG_ColorChoice discrimination in the color descriptors.
 */
export const createColorElement = (color: SolidFillOptions): string => emitColorChoice(color);

/**
 * Creates a solid fill element as an XML string.
 *
 * Specifies a solid color fill using any supported color type.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_SolidColorFillProperties">
 *   <xsd:sequence>
 *     <xsd:group ref="EG_ColorChoice" minOccurs="0"/>
 *     <xsd:group ref="EG_EffectProperties" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // RGB solid fill
 * const fill = createSolidFill({ value: "FF0000" });
 * // Scheme solid fill with tint
 * const schemeFill = createSolidFill({
 *   value: SchemeColor.ACCENT1, transforms: { tint: 40 },
 * });
 * // HSL solid fill
 * const hslFill = createSolidFill({ hue: 120, saturation: 100, luminance: 50 });
 * ```
 */
export const createSolidFill = (options: SolidFillOptions): string =>
  element("a:solidFill", undefined, [createColorElement(options)]);
