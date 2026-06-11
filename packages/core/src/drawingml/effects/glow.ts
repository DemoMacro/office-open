/**
 * Glow effect for DrawingML shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_GlowEffect
 *
 * @module
 */
import { element } from "@office-open/xml";

import { createColorElement } from "../color/solid-fill";
import type { SolidFillOptions } from "../color/solid-fill";

/**
 * Options for glow effect.
 */
export interface GlowEffectOptions {
  /** Glow radius in EMUs */
  radius?: number;
  /** Glow color */
  color: SolidFillOptions;
}

/**
 * Creates a glow effect element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_GlowEffect">
 *   <xsd:sequence>
 *     <xsd:group ref="EG_ColorChoice" minOccurs="1" maxOccurs="1"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="rad" type="ST_PositiveCoordinate" use="optional" default="0"/>
 * </xsd:complexType>
 * ```
 */
export const createGlowEffect = (options: GlowEffectOptions): string => {
  const colorChild = createColorElement(options.color);
  if (options.radius === undefined) {
    return element("a:glow", undefined, [colorChild]);
  }

  return element("a:glow", { rad: options.radius }, [colorChild]);
};
