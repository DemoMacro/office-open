/**
 * Table layout module for WordprocessingML documents.
 *
 * This module provides table layout algorithm settings.
 *
 * @module
 */

/**
 * Table layout algorithm types.
 *
 * Specifies how the table width is calculated.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_TblLayoutType">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="fixed"/>
 *     <xsd:enumeration value="autofit"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 *
 * @publicApi
 */
export const TableLayoutType = {
  /** Auto-fit layout - column widths are adjusted based on content */
  AUTOFIT: "autofit",
  /** Fixed layout - column widths are fixed as specified */
  FIXED: "fixed",
} as const;
