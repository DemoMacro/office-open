/**
 * Line end (arrow) properties for DrawingML outlines.
 *
 * This module provides support for line end markers (arrows) on shape outlines.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_LineEndProperties
 *
 * @module
 */
import { element } from "@office-open/xml";

import { xsdLineEndSize } from "../../util/mappings";

/**
 * Line end types (arrow head styles).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_LineEndType">
 *   <xsd:restriction base="xsd:token">
 *     <xsd:enumeration value="none"/>
 *     <xsd:enumeration value="triangle"/>
 *     <xsd:enumeration value="stealth"/>
 *     <xsd:enumeration value="diamond"/>
 *     <xsd:enumeration value="oval"/>
 *     <xsd:enumeration value="arrow"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 *
 * @publicApi
 */
export const LineEndType = {
  /** No line end */
  NONE: "none",
  /** Triangle arrow */
  TRIANGLE: "triangle",
  /** Stealth arrow (filled triangle) */
  STEALTH: "stealth",
  /** Diamond shape */
  DIAMOND: "diamond",
  /** Oval shape */
  OVAL: "oval",
  /** Simple arrow */
  ARROW: "arrow",
} as const;

/**
 * Line end width options.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_LineEndWidth">
 *   <xsd:restriction base="xsd:token">
 *     <xsd:enumeration value="sm"/>
 *     <xsd:enumeration value="med"/>
 *     <xsd:enumeration value="lg"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 *
 * @publicApi
 */
export const LineEndWidth = {
  /** Small width */
  SMALL: "small",
  /** Medium width */
  MEDIUM: "medium",
  /** Large width */
  LARGE: "large",
} as const;

/**
 * Line end length options.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_LineEndLength">
 *   <xsd:restriction base="xsd:token">
 *     <xsd:enumeration value="sm"/>
 *     <xsd:enumeration value="med"/>
 *     <xsd:enumeration value="lg"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 *
 * @publicApi
 */
export const LineEndLength = {
  /** Small length */
  SMALL: "small",
  /** MEDIUM length */
  MEDIUM: "medium",
  /** Large length */
  LARGE: "large",
} as const;

/**
 * Options for line end (arrow) properties.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_LineEndProperties">
 *   <xsd:attribute name="type" type="ST_LineEndType" use="optional" default="none"/>
 *   <xsd:attribute name="w" type="ST_LineEndWidth" use="optional"/>
 *   <xsd:attribute name="len" type="ST_LineEndLength" use="optional"/>
 * </xsd:complexType>
 * ```
 */
export interface LineEndOptions {
  /** Arrow/head type */
  type: (typeof LineEndType)[keyof typeof LineEndType];
  /** Arrow width */
  width?: (typeof LineEndWidth)[keyof typeof LineEndWidth];
  /** Arrow length */
  length?: (typeof LineEndLength)[keyof typeof LineEndLength];
}

/**
 * Creates a line end element (a:headEnd or a:tailEnd).
 *
 * @example
 * ```typescript
 * // Stealth arrow at start, medium size
 * createLineEnd("a:headEnd", { type: "STEALTH", width: "MEDIUM", length: "MEDIUM" });
 * // Triangle arrow at end
 * createLineEnd("a:tailEnd", { type: "TRIANGLE" });
 * ```
 */
export const createLineEnd = (name: string, options: LineEndOptions): string =>
  element(name, {
    type: options.type,
    w: options.width ? xsdLineEndSize.to(options.width) : undefined,
    len: options.length ? xsdLineEndSize.to(options.length) : undefined,
  });
