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
export type RectAlignment =
  | "topLeft"
  | "top"
  | "topRight"
  | "left"
  | "center"
  | "right"
  | "bottomLeft"
  | "bottom"
  | "bottomRight";

/**
 * Options for outer shadow effect.
 */
export interface OuterShadowEffectOptions {
  /** Blur radius in EMUs */
  blurRadius?: number;
  /** Distance from shape in EMUs */
  distance?: number;
  /** Direction angle in 60,000ths of a degree (will take degrees in the angle batch) */
  direction?: number;
  /** Horizontal scale as integer percent (100 = 100%) */
  scaleX?: number;
  /** Vertical scale as integer percent */
  scaleY?: number;
  /** Horizontal skew angle in 60,000ths of a degree (will take degrees in the angle batch) */
  skewX?: number;
  /** Vertical skew angle */
  skewY?: number;
  /** Shadow alignment */
  alignment?: RectAlignment;
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
  if (options.scaleX !== undefined) attrs.sx = options.scaleX * 1000;
  if (options.scaleY !== undefined) attrs.sy = options.scaleY * 1000;
  if (options.skewX !== undefined) attrs.kx = options.skewX;
  if (options.skewY !== undefined) attrs.ky = options.skewY;
  if (options.alignment !== undefined) attrs.algn = xsdRectAlignment.to(options.alignment);
  if (options.rotWithShape === false) attrs.rotWithShape = 0;

  return element("a:outerShdw", attrs, [createColorElement(options.color)]);
};
