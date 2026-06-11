/**
 * Table cell spacing module for WordprocessingML documents.
 *
 * This module provides cell spacing settings for tables, controlling
 * the space between cells in a table.
 *
 * Reference: http://officeopenxml.com/WPtableCellSpacing.php
 *
 * @module
 */
import type { Percentage, UniversalMeasure } from "@office-open/core";

/**
 * Cell spacing measurement types.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_TblCellSpacing">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="nil"/>
 *     <xsd:enumeration value="dxa"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 */
export const CellSpacingType = {
  /** Value is in twentieths of a point */
  DXA: "dxa",
  /** No (empty) value. */
  NIL: "nil",
} as const;

/**
 * Properties for table cell spacing.
 */
export interface TableCellSpacingProperties {
  /** The spacing value (in twips, percentage, or universal measure) */
  value: number | Percentage | UniversalMeasure;
  /** The type of measurement (defaults to DXA/twips) */
  type?: (typeof CellSpacingType)[keyof typeof CellSpacingType];
}
