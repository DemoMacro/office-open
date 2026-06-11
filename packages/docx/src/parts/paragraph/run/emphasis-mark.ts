/**
 * Emphasis mark module for WordprocessingML run properties.
 *
 * This module provides support for East Asian emphasis marks, which are
 * characters placed above or below text to emphasize it.
 *
 * Reference: http://officeopenxml.com/WPrun.php
 *
 * @module
 */

/**
 * Emphasis mark types.
 *
 * Defines the types of emphasis marks that can be applied to text.
 * Emphasis marks are commonly used in East Asian typography.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_Em">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="none"/>
 *     <xsd:enumeration value="dot"/>
 *     <xsd:enumeration value="comma"/>
 *     <xsd:enumeration value="circle"/>
 *     <xsd:enumeration value="underDot"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 *
 * @publicApi
 */
export const EmphasisMarkType = {
  /** No emphasis mark */
  NONE: "none",
  /** Comma emphasis mark */
  COMMA: "comma",
  /** Circle emphasis mark */
  CIRCLE: "circle",
  /** Dot emphasis mark */
  DOT: "dot",
  /** Under dot emphasis mark */
  UNDER_DOT: "underDot",
} as const;
