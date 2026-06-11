/**
 * Inner shadow effect for DrawingML shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_InnerShadowEffect
 *
 * @module
 */
import { element } from "@office-open/xml";

import { createColorElement } from "../color/solid-fill";
import type { SolidFillOptions } from "../color/solid-fill";

/**
 * Options for inner shadow effect.
 */
export interface InnerShadowEffectOptions {
  /** Blur radius in EMUs */
  blurRadius?: number;
  /** Distance from shape edge in EMUs */
  distance?: number;
  /** Direction angle in 60,000ths of a degree */
  direction?: number;
  /** Shadow color */
  color: SolidFillOptions;
}

/**
 * Creates an inner shadow effect element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_InnerShadowEffect">
 *   <xsd:sequence>
 *     <xsd:group ref="EG_ColorChoice" minOccurs="1" maxOccurs="1"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="blurRad" type="ST_PositiveCoordinate" default="0"/>
 *   <xsd:attribute name="dist" type="ST_PositiveCoordinate" default="0"/>
 *   <xsd:attribute name="dir" type="ST_PositiveFixedAngle" default="0"/>
 * </xsd:complexType>
 * ```
 */
export const createInnerShadowEffect = (options: InnerShadowEffectOptions): string => {
  const colorChild = createColorElement(options.color);

  const attrs: Record<string, number> = {};
  if (options.blurRadius !== undefined) attrs.blurRad = options.blurRadius;
  if (options.distance !== undefined) attrs.dist = options.distance;
  if (options.direction !== undefined) attrs.dir = options.direction;

  const hasAttributes = Object.keys(attrs).length > 0;

  return element("a:innerShdw", hasAttributes ? attrs : undefined, [colorChild]);
};
