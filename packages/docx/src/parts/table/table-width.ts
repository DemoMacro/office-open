/**
 * Table width module for WordprocessingML documents.
 *
 * This module provides width specifications for tables and cells.
 *
 * Reference: http://officeopenxml.com/WPtableWidth.php
 *
 * @module
 */
import type { Percentage, UniversalMeasure } from "@office-open/core";

/**
 * Width type values for tables and cells.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_TblWidth">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="nil"/>
 *     <xsd:enumeration value="pct"/>
 *     <xsd:enumeration value="dxa"/>
 *     <xsd:enumeration value="auto"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 *
 * @publicApi
 */
export const WidthType = {
  /** Auto. */
  AUTO: "auto",
  /** Value is in twentieths of a point */
  DXA: "dxa",
  /** No (empty) value. */
  NIL: "nil",
  /** Value is in percentage. */
  PERCENTAGE: "pct",
} as const;

/**
 * Properties for specifying table or cell width.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_TblWidth">
 *   <xsd:attribute name="w" type="ST_MeasurementOrPercent"/>
 *   <xsd:attribute name="type" type="ST_TblWidth"/>
 * </xsd:complexType>
 * ```
 */
export interface TableWidthProperties {
  size: number | Percentage | UniversalMeasure;
  type?: (typeof WidthType)[keyof typeof WidthType];
}
