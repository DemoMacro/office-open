/**
 * Gradient fill element for DrawingML shapes.
 *
 * This module provides gradient fill support with linear and path shading.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_GradientFillProperties
 *
 * @module
 */
import { element } from "@office-open/xml";

import type { SolidFillOptions } from "../color/solid-fill";
import { createColorElement } from "../color/solid-fill";

/**
 * Gradient stop position (0-100000, representing 0%-100%).
 */
export interface GradientStop {
  /** Position of the color stop (0-100000) */
  position: number;
  /** Color at this stop */
  color: SolidFillOptions;
}

/**
 * Path shade type for radial gradients.
 */
export const PathShadeType = {
  /** Follow shape path */
  SHAPE: "shape",
  /** Circular gradient */
  CIRCLE: "circle",
  /** Rectangular gradient */
  RECT: "rect",
} as const;

/**
 * Tile flip mode for gradient fill.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_TileFlipMode">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="none"/>
 *     <xsd:enumeration value="x"/>
 *     <xsd:enumeration value="y"/>
 *     <xsd:enumeration value="xy"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 */
export const TileFlipMode = {
  /** No flip */
  NONE: "none",
  /** Flip horizontally */
  X: "x",
  /** Flip vertically */
  Y: "y",
  /** Flip both horizontally and vertically */
  XY: "xy",
} as const;

/**
 * Options for linear gradient shading.
 */
export interface LinearShadeOptions {
  /** Angle in 60,000ths of a degree (e.g., 5400000 = 90°) */
  angle?: number;
  /** Whether the angle scales with the shape */
  scaled?: boolean;
}

/**
 * Relative rectangle (CT_RelativeRect) with percentage offsets.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_RelativeRect">
 *   <xsd:attribute name="l" type="ST_Percentage" default="0%"/>
 *   <xsd:attribute name="t" type="ST_Percentage" default="0%"/>
 *   <xsd:attribute name="r" type="ST_Percentage" default="0%"/>
 *   <xsd:attribute name="b" type="ST_Percentage" default="0%"/>
 * </xsd:complexType>
 * ```
 */
export interface RelativeRect {
  /** Left offset percentage (e.g., "0%") */
  left?: string;
  /** Top offset percentage (e.g., "0%") */
  top?: string;
  /** Right offset percentage (e.g., "0%") */
  right?: string;
  /** Bottom offset percentage (e.g., "0%") */
  bottom?: string;
}

/**
 * Options for path (radial) gradient shading.
 */
export interface PathShadeOptions {
  /** Path type */
  path?: (typeof PathShadeType)[keyof typeof PathShadeType];
  /**
   * Fill-to rectangle for path gradient.
   *
   * Defines the rectangle to which the gradient fills.
   */
  fillToRect?: RelativeRect;
}

/**
 * Gradient shade options (linear or path).
 */
export type GradientShadeOptions = LinearShadeOptions | PathShadeOptions;

/**
 * Options for gradient fill.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_GradientFillProperties">
 *   <xsd:sequence>
 *     <xsd:element name="gsLst" type="CT_GradientStopList" minOccurs="0"/>
 *     <xsd:group ref="EG_ShadeProperties" minOccurs="0"/>
 *     <xsd:element name="tileRect" type="CT_RelativeRect" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="flip" type="ST_TileFlipMode" use="optional"/>
 *   <xsd:attribute name="rotWithShape" type="xsd:boolean" use="optional"/>
 * </xsd:complexType>
 * ```
 */
export interface GradientFillOptions {
  /** Gradient color stops (minimum 2) */
  stops: readonly GradientStop[];
  /** Shade type (linear or path) */
  shade?: GradientShadeOptions;
  /**
   * Tile flip mode.
   *
   * Controls how the gradient is flipped when tiled.
   */
  flip?: (typeof TileFlipMode)[keyof typeof TileFlipMode];
  /**
   * Tile rectangle for gradient tiling.
   *
   * Defines the rectangle used for gradient tiling.
   */
  tileRect?: RelativeRect;
  /** Whether gradient rotates with the shape */
  rotateWithShape?: boolean;
}

/**
 * Creates a gradient stop element (a:gs).
 *
 * @example
 * ```typescript
 * createGradientStop({ position: 0, color: { value: "FF0000" } });
 * createGradientStop({ position: 100000, color: { value: "0000FF" } });
 * ```
 */
export const createGradientStop = (stop: GradientStop): string =>
  element("a:gs", { pos: stop.position }, [createColorElement(stop.color)]);

/**
 * Creates a relative rect element.
 */
const createRelativeRect = (name: string, rect?: RelativeRect): string =>
  element(name, {
    l: rect?.left,
    t: rect?.top,
    r: rect?.right,
    b: rect?.bottom,
  });

/**
 * Creates the shade element (a:lin or a:path).
 */
const createShadeElement = (shade: GradientShadeOptions): string => {
  if ("angle" in shade) {
    return element("a:lin", {
      ang: shade.angle,
      scaled: shade.scaled,
    });
  }
  const pathShade = shade as PathShadeOptions;
  const children: string[] = [];

  if (pathShade.fillToRect) {
    children.push(createRelativeRect("a:fillToRect", pathShade.fillToRect));
  }

  return element("a:path", { path: pathShade.path }, children);
};

/**
 * Creates a gradient fill element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_GradientFillProperties">
 *   <xsd:sequence>
 *     <xsd:element name="gsLst" type="CT_GradientStopList" minOccurs="0"/>
 *     <xsd:group ref="EG_ShadeProperties" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="rotWithShape" type="xsd:boolean" use="optional"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Linear gradient from red to blue
 * createGradientFill({
 *   stops: [
 *     { position: 0, color: { value: "FF0000" } },
 *     { position: 100000, color: { value: "0000FF" } },
 *   ],
 *   shade: { angle: 5400000 },
 * });
 * ```
 */
export const createGradientFill = (options: GradientFillOptions): string => {
  const children: string[] = [];

  // Gradient stop list
  const stopElements = options.stops.map(createGradientStop);
  children.push(element("a:gsLst", undefined, stopElements));

  // Shade properties (a:lin or a:path)
  if (options.shade) {
    children.push(createShadeElement(options.shade));
  }

  // Tile rectangle
  if (options.tileRect) {
    children.push(createRelativeRect("a:tileRect", options.tileRect));
  }

  return element(
    "a:gradFill",
    {
      flip: options.flip,
      rotWithShape: options.rotateWithShape,
    },
    children,
  );
};
