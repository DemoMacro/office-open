/**
 * Vertical alignment module for WordprocessingML documents.
 *
 * This module provides vertical alignment options for table cells and sections.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_VerticalJc">
 *   <xsd:attribute name="val" type="ST_VerticalJc" use="required"/>
 * </xsd:complexType>
 *
 * <xsd:simpleType name="ST_VerticalJc">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="both"/>
 *     <xsd:enumeration value="top"/>
 *     <xsd:enumeration value="center"/>
 *     <xsd:enumeration value="bottom"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 *
 * @module
 */
/**
 * Enumeration for table-cell vertical alignment. Only `top`, `center`, `bottom`
 * are valid according to ECMA-376 (§17.18.87 ST_VerticalJc within `<w:tcPr>`).
 *
 * @publicApi
 */
export const VerticalAlignTable = {
  BOTTOM: "bottom",
  CENTER: "center",
  TOP: "top",
} as const;

/**
 * Enumeration for section (<w:sectPr>) vertical alignment. Adds `both` on top of
 * the table-cell set (§17.18.87 ST_VerticalJc within <w:sectPr>).
 *
 * @publicApi
 */
export const VerticalAlignSection = {
  ...VerticalAlignTable,
  BOTH: "both",
} as const;

/** Vertical alignment of content within a table cell (ST_VerticalJc). */
export type TableVerticalAlign = (typeof VerticalAlignTable)[keyof typeof VerticalAlignTable];

/** Vertical alignment of content on the page (ST_VerticalJc in w:sectPr); "both" = justified top-to-bottom. */
export type SectionVerticalAlign = (typeof VerticalAlignSection)[keyof typeof VerticalAlignSection];
