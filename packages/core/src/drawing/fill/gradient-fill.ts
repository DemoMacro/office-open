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

import { stripColorHashPrefix } from "../../util/values";
import type { TileFlipMode } from "../blip/tile";
import type { SolidFillOptions } from "../color/solid-fill";
import { createColorElement } from "../color/solid-fill";
import type { GradientStopOptions } from "./fill-options";

// Single home for the shared ST_TileFlipMode token set is blip/tile.ts;
// re-exported here for gradient-fill consumers.
export type { TileFlipMode } from "../blip/tile";

/** Narrow a stop color to SolidFillOptions for EG_ColorChoice emission. */
export const toSolidColor = (color: GradientStopOptions["color"]): SolidFillOptions =>
  typeof color === "string" ? ({ value: stripColorHashPrefix(color) } as SolidFillOptions) : color;

/**
 * Path shade type for radial gradients.
 */
export type PathShade = "shape" | "circle" | "rect";

/**
 * Options for linear gradient shading.
 */
export interface LinearShadeOptions {
  /** Angle in degrees (e.g., 90 = 90°). */
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
  path?: PathShade;
  /**
   * Fill-to rectangle for path gradient.
   *
   * Defines the rectangle to which the gradient fills.
   */
  fillToRectangle?: RelativeRect;
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
  /** Gradient color stops (minimum 2); color accepts a hex string or SolidFillOptions */
  stops: readonly GradientStopOptions[];
  /** Shade type (linear or path) */
  shade?: GradientShadeOptions;
  /**
   * Tile flip mode.
   *
   * Controls how the gradient is flipped when tiled.
   */
  flip?: TileFlipMode;
  /**
   * Tile rectangle for gradient tiling.
   *
   * Defines the rectangle used for gradient tiling.
   */
  tileRectangle?: RelativeRect;
  /** Whether gradient rotates with the shape */
  rotateWithShape?: boolean;
}

/**
 * Creates a gradient stop element (a:gs).
 *
 * @example
 * ```typescript
 * createGradientStop({ position: 0, color: "FF0000" });
 * createGradientStop({ position: 100, color: { value: "0000FF" } });
 * ```
 */
export const createGradientStop = (stop: GradientStopOptions): string =>
  element("a:gs", { pos: Math.round(stop.position * 1000) }, [
    createColorElement(toSolidColor(stop.color)),
  ]);

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
      ang: shade.angle !== undefined ? Math.round(shade.angle * 60000) : undefined,
      scaled: shade.scaled,
    });
  }
  const pathShade = shade as PathShadeOptions;
  const children: string[] = [];

  if (pathShade.fillToRectangle) {
    children.push(createRelativeRect("a:fillToRect", pathShade.fillToRectangle));
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
 *     { position: 100, color: { value: "0000FF" } },
 *   ],
 *   shade: { angle: 90 },
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
  if (options.tileRectangle) {
    children.push(createRelativeRect("a:tileRect", options.tileRectangle));
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
