/**
 * Table row height module for WordprocessingML documents.
 *
 * This module provides row height configuration including rules for how height should be applied.
 *
 * Reference: http://officeopenxml.com/WPtableRow.php
 *
 * @module
 */

/**
 * Height rules for table rows.
 *
 * Specifies how the height value should be interpreted.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_HeightRule">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="auto"/>
 *     <xsd:enumeration value="exact"/>
 *     <xsd:enumeration value="atLeast"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 *
 * @publicApi
 */
export const HeightRule = {
  /** Height is determined based on the content, so value is ignored. */
  AUTO: "auto",
  /** At least the value specified */
  ATLEAST: "atLeast",
  /** Exactly the value specified */
  EXACT: "exact",
} as const;
