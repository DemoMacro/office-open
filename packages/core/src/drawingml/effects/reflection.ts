/**
 * Reflection effect for DrawingML shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_ReflectionEffect
 *
 * @module
 */
import { element } from "@office-open/xml";

import { xsdRectAlignment } from "../../util/mappings";

/**
 * Options for reflection effect.
 *
 * All properties are optional with XSD defaults.
 */
export interface ReflectionEffectOptions {
  /** Blur radius in EMUs */
  blurRadius?: number;
  /** Start opacity (fixed percentage, e.g., 100000 = 100%) */
  startAlpha?: number;
  /** Start position (fixed percentage) */
  startPosition?: number;
  /** End opacity (fixed percentage) */
  endAlpha?: number;
  /** End position (fixed percentage) */
  endPosition?: number;
  /** Distance from shape in EMUs */
  distance?: number;
  /** Direction angle in 60,000ths of a degree */
  direction?: number;
  /** Fade direction angle in 60,000ths of a degree */
  fadeDirection?: number;
  /** Horizontal scale percentage */
  scaleX?: number;
  /** Vertical scale percentage */
  scaleY?: number;
  /** Horizontal skew angle */
  skewX?: number;
  /** Vertical skew angle */
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

  if (options.blurRadius !== undefined) attrs.blurRad = options.blurRadius;
  if (options.startAlpha !== undefined) attrs.stA = options.startAlpha;
  if (options.startPosition !== undefined) attrs.stPos = options.startPosition;
  if (options.endAlpha !== undefined) attrs.endA = options.endAlpha;
  if (options.endPosition !== undefined) attrs.endPos = options.endPosition;
  if (options.distance !== undefined) attrs.dist = options.distance;
  if (options.direction !== undefined) attrs.dir = options.direction;
  if (options.fadeDirection !== undefined) attrs.fadeDir = options.fadeDirection;
  if (options.scaleX !== undefined) attrs.sx = options.scaleX;
  if (options.scaleY !== undefined) attrs.sy = options.scaleY;
  if (options.skewX !== undefined) attrs.kx = options.skewX;
  if (options.skewY !== undefined) attrs.ky = options.skewY;
  if (options.alignment !== undefined) attrs.algn = xsdRectAlignment.to(options.alignment);
  if (options.rotWithShape === false) attrs.rotWithShape = 0;

  return element("a:reflection", attrs);
};
