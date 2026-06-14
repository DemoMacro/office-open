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
import type { TableWidthProperties } from "@parts/table/table-width";

/**
 * Options for configuring table cell margins (CT_TblCellMar).
 *
 * Each side is an independent CT_TblWidth ({ size, type }), mirroring
 * TableWidthProperties / TableCellSpacingProperties. This is faithful to the
 * XSD, which models CT_TblCellMar as a sequence of six CT_TblWidth elements —
 * every side carries its own width and unit type (no shared `marginUnitType`).
 */
export interface TableCellMarginOptions {
  /** Top cell margin. */
  top?: TableWidthProperties;
  /** Start cell margin (logical side; mirrors `left` in LTR, `right` in RTL). */
  start?: TableWidthProperties;
  /** Left cell margin (physical side). */
  left?: TableWidthProperties;
  /** Bottom cell margin. */
  bottom?: TableWidthProperties;
  /** End cell margin (logical side; mirrors `right` in LTR, `left` in RTL). */
  end?: TableWidthProperties;
  /** Right cell margin (physical side). */
  right?: TableWidthProperties;
}
