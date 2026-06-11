/**
 * Outer shadow effect for DrawingML shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_OuterShadowEffect
 *
 * @module
 */
import { element } from "@office-open/xml";

import { xsdRectAlignment } from "../../util/mappings";
import { createColorElement } from "../color/solid-fill";
import type { SolidFillOptions } from "../color/solid-fill";

/**
 * Rectangle alignment for shadow positioning.
 */
export const RectAlignment = {
  TOP_LEFT: "topLeft",
  TOP: "top",
  TOP_RIGHT: "topRight",
  LEFT: "left",
  CENTER: "center",
  RIGHT: "right",
  BOTTOM_LEFT: "bottomLeft",
  BOTTOM: "bottom",
  BOTTOM_RIGHT: "bottomRight",
} as const;

/**
 * Options for outer shadow effect.
 */
export interface OuterShadowEffectOptions {
  /** Blur radius in EMUs */
  blurRadius?: number;
  /** Distance from shape in EMUs */
  distance?: number;
  /** Direction angle in 60,000ths of a degree */
  direction?: number;
  /** Horizontal scale percentage (e.g., 100000 = 100%) */
  scaleX?: number;
  /** Vertical scale percentage */
  scaleY?: number;
  /** Horizontal skew angle in 60,000ths of a degree */
  skewX?: number;
  /** Vertical skew angle */
  skewY?: number;
  /** Shadow alignment */
  alignment?: (typeof RectAlignment)[keyof typeof RectAlignment];
  /** Whether shadow rotates with shape */
  rotWithShape?: boolean;
  /** Shadow color */
  color: SolidFillOptions;
}

/**
 * Creates an outer shadow effect element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_OuterShadowEffect">
 *   <xsd:sequence>
 *     <xsd:group ref="EG_ColorChoice" minOccurs="1" maxOccurs="1"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="blurRad" type="ST_PositiveCoordinate" default="0"/>
 *   <xsd:attribute name="dist" type="ST_PositiveCoordinate" default="0"/>
 *   <xsd:attribute name="dir" type="ST_PositiveFixedAngle" default="0"/>
 *   <xsd:attribute name="sx" type="ST_Percentage" default="100%"/>
 *   <xsd:attribute name="sy" type="ST_Percentage" default="100%"/>
 *   <xsd:attribute name="kx" type="ST_FixedAngle" default="0"/>
 *   <xsd:attribute name="ky" type="ST_FixedAngle" default="0"/>
 *   <xsd:attribute name="algn" type="ST_RectAlignment" default="b"/>
 *   <xsd:attribute name="rotWithShape" type="xsd:boolean" default="true"/>
 * </xsd:complexType>
 * ```
 */
export const createOuterShadowEffect = (options: OuterShadowEffectOptions): string => {
  const attrs: Record<string, string | number> = {};

  if (options.blurRadius !== undefined) attrs.blurRad = options.blurRadius;
  if (options.distance !== undefined) attrs.dist = options.distance;
  if (options.direction !== undefined) attrs.dir = options.direction;
  if (options.scaleX !== undefined) attrs.sx = options.scaleX;
  if (options.scaleY !== undefined) attrs.sy = options.scaleY;
  if (options.skewX !== undefined) attrs.kx = options.skewX;
  if (options.skewY !== undefined) attrs.ky = options.skewY;
  if (options.alignment !== undefined) attrs.algn = xsdRectAlignment.to(options.alignment);
  if (options.rotWithShape === false) attrs.rotWithShape = 0;

  return element("a:outerShdw", attrs, [createColorElement(options.color)]);
};
