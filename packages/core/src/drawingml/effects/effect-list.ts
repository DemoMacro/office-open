/**
 * Effect list container for DrawingML shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_EffectList, EG_EffectProperties
 *
 * @module
 */
import { element } from "@office-open/xml";

import type { FillOverlayEffectOptions } from "./fill-overlay";
import { createFillOverlayEffect } from "./fill-overlay";
import type { GlowEffectOptions } from "./glow";
import { createGlowEffect } from "./glow";
import type { InnerShadowEffectOptions } from "./inner-shadow";
import { createInnerShadowEffect } from "./inner-shadow";
import type { OuterShadowEffectOptions } from "./outer-shadow";
import { createOuterShadowEffect } from "./outer-shadow";
import type { PresetShadowEffectOptions } from "./preset-shadow";
import { createPresetShadowEffect } from "./preset-shadow";
import type { ReflectionEffectOptions } from "./reflection";
import { createReflectionEffect } from "./reflection";
import { createSoftEdgeEffect } from "./soft-edge";

/**
 * Effect extent calculation result.
 * Represents the additional space needed on each edge (in EMUs).
 */
export interface EffectExtent {
  /** Left edge extension */
  l: number;
  /** Top edge extension */
  t: number;
  /** Right edge extension */
  r: number;
  /** Bottom edge extension */
  b: number;
}

/**
 * Calculates the effect extent for a given set of effects.
 *
 * The effectExtent specifies the additional extent which shall be added to each edge
 * of the image (top, bottom, left, right) in order to compensate for any drawing effects
 * applied to the DrawingML object.
 *
 * Reference: CT_EffectExtent in dml-wordprocessingDrawing.xsd
 *
 * @param options - Effect list options
 * @returns The calculated effect extent in EMUs
 */
export const calculateEffectExtent = (options?: EffectListOptions): EffectExtent => {
  if (!options) {
    return { l: 0, t: 0, r: 0, b: 0 };
  }

  let l = 0,
    t = 0,
    r = 0,
    b = 0;

  // Outer shadow: distance + blurRadius extends in the shadow direction
  if (options.outerShadow) {
    const dist = options.outerShadow.distance ?? 0;
    const blur = options.outerShadow.blurRadius ?? 0;
    const dir = options.outerShadow.direction ?? 0;

    // Direction is in 60,000ths of a degree, convert to radians
    const radians = (dir / 60000) * (Math.PI / 180);
    const dx = Math.cos(radians) * dist;
    const dy = Math.sin(radians) * dist;

    // Shadow extends in direction + blur radius in all directions
    const extend = Math.abs(dx) + blur;
    const extendY = Math.abs(dy) + blur;

    if (dx > 0) r = Math.max(r, extend);
    else l = Math.max(l, extend);

    if (dy > 0) b = Math.max(b, extendY);
    else t = Math.max(t, extendY);
  }

  // Glow: radius extends equally in all directions
  if (options.glow) {
    const radius = options.glow.radius ?? 0;
    l = Math.max(l, radius);
    t = Math.max(t, radius);
    r = Math.max(r, radius);
    b = Math.max(b, radius);
  }

  // Reflection: typically extends downward
  if (options.reflection && options.reflection !== true) {
    const dist = options.reflection.distance ?? 0;
    const blur = options.reflection.blurRadius ?? 0;
    // Reflection typically extends downward
    b = Math.max(b, dist + blur);
  }

  // Soft edge: radius extends equally in all directions
  if (options.softEdge !== undefined) {
    l = Math.max(l, options.softEdge);
    t = Math.max(t, options.softEdge);
    r = Math.max(r, options.softEdge);
    b = Math.max(b, options.softEdge);
  }

  return { l, t, r, b };
};

/**
 * Blur effect options.
 */
export interface BlurEffectOptions {
  /** Blur radius in EMUs */
  radius?: number;
  /** Whether to grow the shape boundary */
  grow?: boolean;
}

/**
 * Options for the effect list container.
 *
 * Each property corresponds to an optional child effect element.
 * Only specified effects will be included in the output.
 */
export interface EffectListOptions {
  /** Blur effect */
  blur?: BlurEffectOptions;
  /** Fill overlay effect */
  fillOverlay?: FillOverlayEffectOptions;
  /** Glow effect */
  glow?: GlowEffectOptions;
  /** Inner shadow effect */
  innerShadow?: InnerShadowEffectOptions;
  /** Outer shadow effect */
  outerShadow?: OuterShadowEffectOptions;
  /** Preset shadow effect */
  presetShadow?: PresetShadowEffectOptions;
  /** Reflection effect (pass object for attributes, or true for defaults) */
  reflection?: ReflectionEffectOptions | true;
  /** Soft edge radius in EMUs */
  softEdge?: number;
}

/**
 * Creates a blur effect element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_BlurEffect">
 *   <xsd:attribute name="rad" type="ST_PositiveCoordinate" default="0"/>
 *   <xsd:attribute name="grow" type="xsd:boolean" default="true"/>
 * </xsd:complexType>
 * ```
 */
const createBlurEffect = (options: BlurEffectOptions): string => {
  const attrs: Record<string, number> = {};
  if (options.radius !== undefined) attrs.rad = options.radius;
  if (options.grow === false) attrs.grow = 0;

  return Object.keys(attrs).length > 0 ? element("a:blur", attrs) : element("a:blur");
};

/**
 * Creates an effect list element (a:effectLst).
 *
 * This is the EG_EffectProperties choice for a flat list of effects.
 * Effects are emitted in XSD order: blur, glow, innerShdw, outerShdw, prstShdw, reflection, softEdge.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_EffectList">
 *   <xsd:sequence>
 *     <xsd:element name="blur" type="CT_BlurEffect" minOccurs="0"/>
 *     <xsd:element name="fillOverlay" type="CT_FillOverlayEffect" minOccurs="0"/>
 *     <xsd:element name="glow" type="CT_GlowEffect" minOccurs="0"/>
 *     <xsd:element name="innerShdw" type="CT_InnerShadowEffect" minOccurs="0"/>
 *     <xsd:element name="outerShdw" type="CT_OuterShadowEffect" minOccurs="0"/>
 *     <xsd:element name="prstShdw" type="CT_PresetShadowEffect" minOccurs="0"/>
 *     <xsd:element name="reflection" type="CT_ReflectionEffect" minOccurs="0"/>
 *     <xsd:element name="softEdge" type="CT_SoftEdgesEffect" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export const createEffectList = (options: EffectListOptions): string => {
  const children: string[] = [];

  if (options.blur) {
    children.push(createBlurEffect(options.blur));
  }
  if (options.fillOverlay) {
    children.push(createFillOverlayEffect(options.fillOverlay));
  }
  if (options.glow) {
    children.push(createGlowEffect(options.glow));
  }
  if (options.innerShadow) {
    children.push(createInnerShadowEffect(options.innerShadow));
  }
  if (options.outerShadow) {
    children.push(createOuterShadowEffect(options.outerShadow));
  }
  if (options.presetShadow) {
    children.push(createPresetShadowEffect(options.presetShadow));
  }
  if (options.reflection) {
    children.push(
      createReflectionEffect(options.reflection === true ? undefined : options.reflection),
    );
  }
  if (options.softEdge !== undefined) {
    children.push(createSoftEdgeEffect(options.softEdge));
  }

  return element("a:effectLst", undefined, children);
};
