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
import { XmlComponent } from "@file/xml-components";
import { twipsMeasureValue } from "@util/values";
import type { PositiveUniversalMeasure } from "@util/values";

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

/**
 * Represents a column definition (col) for a multi-column section layout.
 *
 * This element defines the width and spacing for an individual column when
 * using unequal column widths in a section.
 *
 * Reference: http://officeopenxml.com/WPsectionPr.php
 *
 * @publicApi
 *
 * @example
 * ```typescript
 * // Create a column with specific width and spacing
 * new Column({
 *   width: 3000,
 *   space: 720
 * });
 * ```
 */
export class Column extends XmlComponent {
  public constructor(options: ColumnAttributes) {
    super("w:col");

    this.root.push({
      _attr: {
        "w:w": twipsMeasureValue(options.width),
        ...(options.space !== undefined ? { "w:space": twipsMeasureValue(options.space) } : {}),
      },
    });
  }
}
