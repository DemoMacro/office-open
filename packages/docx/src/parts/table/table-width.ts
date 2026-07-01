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

/**
 * OOXML stores width percentages as fiftieths-of-a-percent integer
 * (`w:w="5000"`, `w:type="pct"` = 100%). The public API exposes them as plain
 * percentages (`size: 100` = 100%); these helpers convert at the stringify/parse
 * boundary so callers never handle the raw 5000 value. A bare number must never
 * be emitted with a "%" suffix — that is a different XSD branch (`s:ST_Percentage`)
 * meaning 5000%, which Word treats as `auto` on `tblW`.
 */

/** Stringify: percentage (`100`, `"50%"`) → OOXML fiftieths integer (`5000`). */
export const widthPctToFiftieths = (
  size: number | Percentage | UniversalMeasure,
): number | Percentage | UniversalMeasure => {
  if (typeof size === "number") return Math.round(size * 50);
  if (size.endsWith("%")) return Math.round(Number(size.slice(0, -1)) * 50);
  return size;
};

/** Parse: OOXML fiftieths (`5000`) → percentage (`100`) when `type` is `"pct"`. */
export const widthFiftiethsToPct = (
  size: number | string | undefined,
  type: string | undefined,
): number | string | undefined => (type === "pct" && typeof size === "number" ? size / 50 : size);
