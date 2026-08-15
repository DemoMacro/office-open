/**
 * Reflection effect for DrawingML shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_ReflectionEffect
 *
 * @module
 */
import { element } from "@office-open/xml";

import { convertToEmu } from "../../util/converters";
import { xsdRectAlignment } from "../../util/mappings";
import type { UniversalMeasure } from "../../util/values";

/**
 * Options for reflection effect.
 *
 * All properties are optional with XSD defaults.
 */
export interface ReflectionEffectOptions {
  /** Blur radius in EMUs (number) or UniversalMeasure (mm/cm/in/pt/pc/pi/px). */
  blurRadius?: number | UniversalMeasure;
  /** Start opacity as integer percent (100 = fully opaque) */
  startAlpha?: number;
  /** Start position as integer percent (0-100) */
  startPosition?: number;
  /** End opacity as integer percent */
  endAlpha?: number;
  /** End position as integer percent (0-100) */
  endPosition?: number;
  /** Distance from shape in EMUs (number) or UniversalMeasure. */
  distance?: number | UniversalMeasure;
  /** Direction angle in degrees (0–360). */
  direction?: number;
  /** Fade direction angle in degrees (0–360). */
  fadeDirection?: number;
  /** Horizontal scale as integer percent (100 = 100%) */
  scaleX?: number;
  /** Vertical scale as integer percent */
  scaleY?: number;
  /** Horizontal skew angle in degrees. */
  skewX?: number;
  /** Vertical skew angle in degrees. */
  skewY?: number;
  /** Alignment */
  alignment?: string;
  /** Whether reflection rotates with shape */
  rotWithShape?: boolean;
}

/**
 * Creates a reflection effect element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_ReflectionEffect">
 *   <xsd:attribute name="blurRad" type="ST_PositiveCoordinate" default="0"/>
 *   <xsd:attribute name="stA" type="ST_PositiveFixedPercentage" default="100%"/>
 *   <xsd:attribute name="stPos" type="ST_PositiveFixedPercentage" default="0%"/>
 *   <xsd:attribute name="endA" type="ST_PositiveFixedPercentage" default="0%"/>
 *   <xsd:attribute name="endPos" type="ST_PositiveFixedPercentage" default="100%"/>
 *   <xsd:attribute name="dist" type="ST_PositiveCoordinate" default="0"/>
 *   <xsd:attribute name="dir" type="ST_PositiveFixedAngle" default="0"/>
 *   <xsd:attribute name="fadeDir" type="ST_PositiveFixedAngle" default="5400000"/>
 *   <xsd:attribute name="sx" type="ST_Percentage" default="100%"/>
 *   <xsd:attribute name="sy" type="ST_Percentage" default="100%"/>
 *   <xsd:attribute name="kx" type="ST_FixedAngle" default="0"/>
 *   <xsd:attribute name="ky" type="ST_FixedAngle" default="0"/>
 *   <xsd:attribute name="algn" type="ST_RectAlignment" default="b"/>
 *   <xsd:attribute name="rotWithShape" type="xsd:boolean" default="true"/>
 * </xsd:complexType>
 * ```
 */
export const createReflectionEffect = (options?: ReflectionEffectOptions): string => {
  if (!options) {
    return "<a:reflection/>";
  }

  const attrs: Record<string, string | number> = {};

  if (options.blurRadius !== undefined) attrs.blurRad = convertToEmu(options.blurRadius);
  if (options.startAlpha !== undefined) attrs.stA = Math.round(options.startAlpha * 1000);
  if (options.startPosition !== undefined) attrs.stPos = Math.round(options.startPosition * 1000);
  if (options.endAlpha !== undefined) attrs.endA = Math.round(options.endAlpha * 1000);
  if (options.endPosition !== undefined) attrs.endPos = Math.round(options.endPosition * 1000);
  if (options.distance !== undefined) attrs.dist = convertToEmu(options.distance);
  if (options.direction !== undefined) attrs.dir = Math.round(options.direction * 60000);
  if (options.fadeDirection !== undefined)
    attrs.fadeDir = Math.round(options.fadeDirection * 60000);
  if (options.scaleX !== undefined) attrs.sx = Math.round(options.scaleX * 1000);
  if (options.scaleY !== undefined) attrs.sy = Math.round(options.scaleY * 1000);
  if (options.skewX !== undefined) attrs.kx = Math.round(options.skewX * 60000);
  if (options.skewY !== undefined) attrs.ky = Math.round(options.skewY * 60000);
  if (options.alignment !== undefined) attrs.algn = xsdRectAlignment.to(options.alignment);
  if (options.rotWithShape === false) attrs.rotWithShape = 0;

  return element("a:reflection", attrs);
};
