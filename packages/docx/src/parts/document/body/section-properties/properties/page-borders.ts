/**
 * Page borders module for WordprocessingML section properties.
 *
 * Defines borders around pages in a document section.
 *
 * Reference: http://officeopenxml.com/WPsectionBorders.php
 *
 * @module
 */
import type { BorderOptions } from "@shared/border";

/**
 * Specifies which pages display the page border.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_PageBorderDisplay">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="allPages"/>
 *     <xsd:enumeration value="firstPage"/>
 *     <xsd:enumeration value="notFirstPage"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 *
 * @publicApi
 */
export const PageBorderDisplay = {
  /** Display border on all pages */
  ALL_PAGES: "allPages",
  /** Display border only on first page */
  FIRST_PAGE: "firstPage",
  /** Display border on all pages except first page */
  NOT_FIRST_PAGE: "notFirstPage",
} as const;

/**
 * Specifies whether page border is positioned relative to page edge or text.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_PageBorderOffset">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="page"/>
 *     <xsd:enumeration value="text"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 *
 * @publicApi
 */
export const PageBorderOffsetFrom = {
  /** Position border relative to page edge */
  PAGE: "page",
  /** Position border relative to text (default) */
  TEXT: "text",
} as const;

/**
 * Specifies z-order of page border relative to intersecting objects.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_PageBorderZOrder">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="front"/>
 *     <xsd:enumeration value="back"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 *
 * @publicApi
 */
export const PageBorderZOrder = {
  /** Display border behind page contents */
  BACK: "back",
  /** Display border in front of page contents (default) */
  FRONT: "front",
} as const;

/**
 * Options for configuring page borders.
 *
 * @property display - Which pages display the border
 * @property offsetFrom - Whether border is positioned relative to page or text
 * @property zOrder - Whether border appears in front or behind page contents
 * @property top - Top border styling
 * @property right - Right border styling
 * @property bottom - Bottom border styling
 * @property left - Left border styling
 */
export interface PageBordersOptions {
  /** Which pages display the border */
  display?: (typeof PageBorderDisplay)[keyof typeof PageBorderDisplay];
  /** Whether border is positioned relative to page or text (default: text) */
  offsetFrom?: (typeof PageBorderOffsetFrom)[keyof typeof PageBorderOffsetFrom];
  /** Whether border appears in front or behind page contents (default: front) */
  zOrder?: (typeof PageBorderZOrder)[keyof typeof PageBorderZOrder];
  /** Top border styling */
  top?: BorderOptions;
  /** Right border styling */
  right?: BorderOptions;
  /** Bottom border styling */
  bottom?: BorderOptions;
  /** Left border styling */
  left?: BorderOptions;
}
