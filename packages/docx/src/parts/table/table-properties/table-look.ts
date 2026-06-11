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
 * Options for configuring table look conditional formatting.
 *
 * These options control which conditional formatting styles are applied
 * to the table.
 */
export interface TableLookOptions {
  /** Apply first row conditional formatting. */
  firstRow?: boolean;
  /** Apply last row conditional formatting. */
  lastRow?: boolean;
  /** Apply first column conditional formatting. */
  firstColumn?: boolean;
  /** Apply last column conditional formatting. */
  lastColumn?: boolean;
  /** Disable horizontal row banding. */
  noHBand?: boolean;
  /** Disable vertical column banding. */
  noVBand?: boolean;
}
