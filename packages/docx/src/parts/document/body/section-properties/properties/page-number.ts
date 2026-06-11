import { decimalNumber } from "@office-open/core";
import { element } from "@office-open/xml";
/**
 * Page number module for WordprocessingML section properties.
 *
 * Defines page numbering format and starting value for document sections.
 *
 * Reference: http://officeopenxml.com/WPSectionPgNumType.php
 *
 * @module
 */
import type { NumberFormat } from "@shared/constants";

/**
 * Specifies the separator character between chapter number and page number.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_ChapterSep">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="hyphen"/>
 *     <xsd:enumeration value="period"/>
 *     <xsd:enumeration value="colon"/>
 *     <xsd:enumeration value="emDash"/>
 *     <xsd:enumeration value="enDash"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 *
 * @publicApi
 */
export const PageNumberSeparator = {
  /** Hyphen separator (-) */
  HYPHEN: "hyphen",
  /** Period separator (.) */
  PERIOD: "period",
  /** Colon separator (:) */
  COLON: "colon",
  /** Em dash separator (—) */
  EM_DASH: "emDash",
  /** En dash separator (–) */
  EN_DASH: "enDash",
} as const;

/**
 * Options for configuring page numbering.
 *
 * @property start - Starting page number for the section
 * @property formatType - Number format (decimal, roman, letter, etc.)
 * @property separator - Separator between chapter and page number
 */
export interface PageNumberTypeAttributes {
  /** Starting page number for the section */
  start?: number;
  /** Number format (decimal, roman, letter, etc., default: decimal) */
  formatType?: (typeof NumberFormat)[keyof typeof NumberFormat];
  /** Separator between chapter and page number (default: hyphen) */
  separator?: (typeof PageNumberSeparator)[keyof typeof PageNumberSeparator];
  /** Heading style ID for chapter numbering */
  chapStyle?: number;
}

/**
 * Creates page numbering settings (pgNumType) for a document section.
 *
 * This element specifies the page numbering format and starting value
 * for all pages in a section.
 *
 * Reference: http://officeopenxml.com/WPSectionPgNumType.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_PageNumber">
 *   <xsd:attribute name="fmt" type="ST_NumberFormat" use="optional" default="decimal"/>
 *   <xsd:attribute name="start" type="ST_DecimalNumber" use="optional"/>
 *   <xsd:attribute name="chapStyle" type="ST_DecimalNumber" use="optional"/>
 *   <xsd:attribute name="chapSep" type="ST_ChapterSep" use="optional" default="hyphen"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Start page numbering at 5 with lowercase roman numerals
 * createPageNumberType({
 *   start: 5,
 *   formatType: NumberFormat.LOWER_ROMAN
 * });
 * ```
 */
export const createPageNumberType = ({
  start,
  formatType,
  separator,
  chapStyle,
}: PageNumberTypeAttributes): string =>
  element("w:pgNumType", {
    "w:chapStyle": chapStyle === undefined ? undefined : decimalNumber(chapStyle),
    "w:fmt": formatType,
    "w:chapSep": separator,
    "w:start": start === undefined ? undefined : decimalNumber(start),
  });
