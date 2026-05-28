/**
 * Reflection effect for DrawingML shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_ReflectionEffect
 *
 * @module
 */
import { BuilderElement } from "../../xml-components";
import { xsdRectAlignment } from "../../xsd-mappings";

/**
 * Options for reflection effect.
 *
 * All properties are optional with XSD defaults.
 */
export interface ReflectionEffectOptions {
  /** Blur radius in EMUs */
  readonly blurRadius?: number;
  /** Start opacity (fixed percentage, e.g., 100000 = 100%) */
  readonly startAlpha?: number;
  /** Start position (fixed percentage) */
  readonly startPosition?: number;
  /** End opacity (fixed percentage) */
  readonly endAlpha?: number;
  /** End position (fixed percentage) */
  readonly endPosition?: number;
  /** Distance from shape in EMUs */
  readonly distance?: number;
  /** Direction angle in 60,000ths of a degree */
  readonly direction?: number;
  /** Fade direction angle in 60,000ths of a degree */
  readonly fadeDirection?: number;
  /** Horizontal scale percentage */
  readonly scaleX?: number;
  /** Vertical scale percentage */
  readonly scaleY?: number;
  /** Horizontal skew angle */
  readonly skewX?: number;
  /** Vertical skew angle */
  readonly skewY?: number;
  /** Alignment */
  readonly alignment?: string;
  /** Whether reflection rotates with shape */
  readonly rotWithShape?: boolean;
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
export const createReflectionEffect = (options?: ReflectionEffectOptions) => {
  if (!options) {
    return new BuilderElement({ name: "a:reflection" });
  }

  const attributes: Record<string, { readonly key: string; readonly value: string | number }> = {};

  if (options.blurRadius !== undefined) {
    attributes.blurRad = { key: "blurRad", value: options.blurRadius };
  }
  if (options.startAlpha !== undefined) {
    attributes.stA = { key: "stA", value: options.startAlpha };
  }
  if (options.startPosition !== undefined) {
    attributes.stPos = { key: "stPos", value: options.startPosition };
  }
  if (options.endAlpha !== undefined) {
    attributes.endA = { key: "endA", value: options.endAlpha };
  }
  if (options.endPosition !== undefined) {
    attributes.endPos = { key: "endPos", value: options.endPosition };
  }
  if (options.distance !== undefined) {
    attributes.dist = { key: "dist", value: options.distance };
  }
  if (options.direction !== undefined) {
    attributes.dir = { key: "dir", value: options.direction };
  }
  if (options.fadeDirection !== undefined) {
    attributes.fadeDir = { key: "fadeDir", value: options.fadeDirection };
  }
  if (options.scaleX !== undefined) {
    attributes.sx = { key: "sx", value: options.scaleX };
  }
  if (options.scaleY !== undefined) {
    attributes.sy = { key: "sy", value: options.scaleY };
  }
  if (options.skewX !== undefined) {
    attributes.kx = { key: "kx", value: options.skewX };
  }
  if (options.skewY !== undefined) {
    attributes.ky = { key: "ky", value: options.skewY };
  }
  if (options.alignment !== undefined) {
    attributes.algn = { key: "algn", value: xsdRectAlignment.to(options.alignment) };
  }
  if (options.rotWithShape === false) {
    attributes.rotWithShape = { key: "rotWithShape", value: 0 };
  }

  return new BuilderElement({
    attributes: attributes as never,
    name: "a:reflection",
  });
};
