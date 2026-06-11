/**
 * Bevel element for DrawingML 3D shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_Bevel
 *
 * @module
 */
import { element } from "@office-open/xml";

/**
 * Bevel preset types (12 variations).
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, ST_BevelPresetType
 */
export const BevelPresetType = {
  RELAXED_INSET: "relaxedInset",
  CIRCLE: "circle",
  SLOPE: "slope",
  CROSS: "cross",
  ANGLE: "angle",
  SOFT_ROUND: "softRound",
  CONVEX: "convex",
  COOL_SLANT: "coolSlant",
  DIVOT: "divot",
  RIBLET: "riblet",
  HARD_EDGE: "hardEdge",
  ART_DECO: "artDeco",
} as const;

/**
 * Options for a bevel element.
 */
export interface BevelOptions {
  /** Bevel width in EMUs (default 76200) */
  w?: number;
  /** Bevel height in EMUs (default 76200) */
  h?: number;
  /** Bevel preset type */
  prst?: (typeof BevelPresetType)[keyof typeof BevelPresetType];
}

/**
 * Creates a bevel element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Bevel">
 *   <xsd:attribute name="w" type="ST_PositiveCoordinate" default="76200"/>
 *   <xsd:attribute name="h" type="ST_PositiveCoordinate" default="76200"/>
 *   <xsd:attribute name="prst" type="ST_BevelPresetType" default="circle"/>
 * </xsd:complexType>
 * ```
 */
export const createBevel = (options?: BevelOptions): string => {
  if (!options) {
    return `<a:bevelT/>`;
  }

  const attrs: Record<string, string | number> = {};

  if (options.w !== undefined) {
    attrs.w = options.w;
  }
  if (options.h !== undefined) {
    attrs.h = options.h;
  }
  if (options.prst !== undefined) {
    attrs.prst = options.prst;
  }

  return element("a:bevelT", attrs);
};

/**
 * Creates a bottom bevel element (a:bevelB).
 */
export const createBottomBevel = (options?: BevelOptions): string => {
  if (!options) {
    return `<a:bevelB/>`;
  }

  const attrs: Record<string, string | number> = {};

  if (options.w !== undefined) {
    attrs.w = options.w;
  }
  if (options.h !== undefined) {
    attrs.h = options.h;
  }
  if (options.prst !== undefined) {
    attrs.prst = options.prst;
  }

  return element("a:bevelB", attrs);
};
