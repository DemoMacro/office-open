import type { ShortHexNumber } from "@office-open/core";
/**
 * Table look module for WordprocessingML documents.
 *
 * Table look specifies conditional formatting settings that determine which
 * special formatting is applied to a table. These settings control whether
 * special formatting is applied to the first row, last row, first column,
 * last column, and whether to display horizontal or vertical banding.
 *
 * Reference: http://officeopenxml.com/WPtblLook.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_TblLook">
 *   <xsd:attribute name="val" type="s:ST_ShortHexNumber"/>
 *   <xsd:attribute name="firstRow" type="s:ST_OnOff"/>
 *   <xsd:attribute name="lastRow" type="s:ST_OnOff"/>
 *   <xsd:attribute name="firstColumn" type="s:ST_OnOff"/>
 *   <xsd:attribute name="lastColumn" type="s:ST_OnOff"/>
 *   <xsd:attribute name="noHBand" type="s:ST_OnOff"/>
 *   <xsd:attribute name="noVBand" type="s:ST_OnOff"/>
 * </xsd:complexType>
 * ```
 *
 * @module
 */
/**
 * Table look conditional formatting flags (firstRow/lastRow/firstCol/lastCol/
 * bandRow/bandCol). The serializer inverts banding (w:noHBand = !bandRow).
 */
export interface TableLookOptions {
  /**
   * Legacy hex bitmask (w:val, ST_ShortHexNumber) — Word 2007 wrote it alone
   * (e.g. "04A0"), Word 2010+ writes the explicit booleans. Round-tripped
   * verbatim as a string so leading zeros survive; authoring needs only the
   * boolean fields.
   */
  val?: ShortHexNumber;
  /** Apply first row conditional formatting. */
  firstRow?: boolean;
  /** Apply last row conditional formatting. */
  lastRow?: boolean;
  /** Apply first column conditional formatting. */
  firstCol?: boolean;
  /** Apply last column conditional formatting. */
  lastCol?: boolean;
  /** Apply horizontal row banding. */
  bandRow?: boolean;
  /** Apply vertical column banding. */
  bandCol?: boolean;
}
