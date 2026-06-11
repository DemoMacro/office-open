/**
 * Column module for WordprocessingML section properties.
 *
 * Defines individual column properties within multi-column layouts.
 *
 * Reference: http://officeopenxml.com/WPsectionPr.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Column">
 *   <xsd:attribute name="w" type="s:ST_TwipsMeasure" use="optional" />
 *   <xsd:attribute name="space" type="s:ST_TwipsMeasure" use="optional" default="0" />
 * </xsd:complexType>
 * ```
 *
 * @module
 */
import type { PositiveUniversalMeasure } from "@office-open/core";

/**
 * Options for configuring individual column properties.
 *
 * @property width - Column width in twips or universal measure
 * @property space - Space after column in twips or universal measure (default: 0)
 */
export interface ColumnAttributes {
  /** Column width in twips or universal measure */
  width: number | PositiveUniversalMeasure;
  /** Space after column in twips or universal measure (default: 0) */
  space?: number | PositiveUniversalMeasure;
}
