/**
 * Inner shadow effect for DrawingML shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_InnerShadowEffect
 *
 * @module
 */
import { element } from "@office-open/xml";

import { convertToEmu } from "../../util/converters";
import type { UniversalMeasure } from "../../util/values";
import { createColorElement } from "../color/solid-fill";
import type { SolidFillOptions } from "../color/solid-fill";

/**
 * Options for inner shadow effect.
 */
export interface InnerShadowEffectOptions {
  /** Blur radius in EMUs (number) or UniversalMeasure (mm/cm/in/pt/pc/pi/px). */
  blurRadius?: number | UniversalMeasure;
  /** Distance from shape edge in EMUs (number) or UniversalMeasure. */
  distance?: number | UniversalMeasure;
  /** Direction angle in degrees (0–360). */
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
  if (options.blurRadius !== undefined) attrs.blurRad = convertToEmu(options.blurRadius);
  if (options.distance !== undefined) attrs.dist = convertToEmu(options.distance);
  if (options.direction !== undefined) attrs.dir = Math.round(options.direction * 60000);

  const hasAttributes = Object.keys(attrs).length > 0;

  return element("a:innerShdw", hasAttributes ? attrs : undefined, [colorChild]);
};
