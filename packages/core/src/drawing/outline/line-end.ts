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
export type LineEndType = "none" | "triangle" | "stealth" | "diamond" | "oval" | "arrow";

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
export type LineEndWidth = "small" | "medium" | "large";

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
export type LineEndLength = "small" | "medium" | "large";

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
  /** Arrow/head type (omitted in source means the OOXML default `none`). */
  type?: LineEndType;
  /** Arrow width */
  width?: LineEndWidth;
  /** Arrow length */
  length?: LineEndLength;
}

/**
 * Creates a line end element (a:headEnd or a:tailEnd).
 *
 * @example
 * ```typescript
 * // Stealth arrow at start, medium size
 * createLineEnd("a:headEnd", { type: "stealth", width: "medium", length: "medium" });
 * // Triangle arrow at end
 * createLineEnd("a:tailEnd", { type: "triangle" });
 * ```
 */
export const createLineEnd = (name: string, options: LineEndOptions): string =>
  element(name, {
    type: options.type,
    w: options.width ? xsdLineEndSize.to(options.width) : undefined,
    len: options.length ? xsdLineEndSize.to(options.length) : undefined,
  });
