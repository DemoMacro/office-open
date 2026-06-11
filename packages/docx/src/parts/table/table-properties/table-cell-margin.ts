/**
 * Table cell margin module for WordprocessingML documents.
 *
 * This module provides cell margin settings for tables and individual cells.
 * Margins define the padding between cell content and cell borders.
 *
 * Reference: http://officeopenxml.com/WPtableCellProperties-Margins.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_TblCellMar">
 *   <xsd:sequence>
 *     <xsd:element name="top" type="CT_TblWidth" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="start" type="CT_TblWidth" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="left" type="CT_TblWidth" minOccurs="0"/>
 *     <xsd:element name="bottom" type="CT_TblWidth" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="end" type="CT_TblWidth" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="right" type="CT_TblWidth" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @module
 */
import type { PositiveUniversalMeasure } from "@office-open/core";
import { WidthType } from "@parts/table";

/**
 * Options for configuring table cell margins.
 *
 * All margin values are specified in the units defined by `marginUnitType`.
 * If `marginUnitType` is not specified, values are in DXA (twentieths of a point).
 */
export interface TableCellMarginOptions {
  /**
   * The unit type for margin values.
   * Defaults to DXA (twentieths of a point) if not specified.
   *
   * @default WidthType.DXA
   */
  marginUnitType?: (typeof WidthType)[keyof typeof WidthType];

  /** Top margin (padding above cell content). */
  top?: number | PositiveUniversalMeasure;

  /** Bottom margin (padding below cell content). */
  bottom?: number | PositiveUniversalMeasure;

  /** Left margin (padding to the left of cell content). */
  left?: number | PositiveUniversalMeasure;

  /** Right margin (padding to the right of cell content). */
  right?: number | PositiveUniversalMeasure;
}
