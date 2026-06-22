/**
 * Table row properties module for WordprocessingML documents.
 *
 * This module provides row-level properties including height and header row settings.
 *
 * Reference: http://officeopenxml.com/WPtableRowProperties.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_TrPrBase">
 *   <xsd:choice maxOccurs="unbounded">
 *     <xsd:element name="cnfStyle" type="CT_Cnf" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="divId" type="CT_DecimalNumber" minOccurs="0"/>
 *     <xsd:element name="gridBefore" type="CT_DecimalNumber" minOccurs="0"/>
 *     <xsd:element name="gridAfter" type="CT_DecimalNumber" minOccurs="0"/>
 *     <xsd:element name="wBefore" type="CT_TblWidth" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="wAfter" type="CT_TblWidth" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="cantSplit" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="trHeight" type="CT_Height" minOccurs="0"/>
 *     <xsd:element name="tblHeader" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="tblCellSpacing" type="CT_TblWidth" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="jc" type="CT_JcTable" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="hidden" type="CT_OnOff" minOccurs="0"/>
 *   </xsd:choice>
 * </xsd:complexType>
 * <xsd:complexType name="CT_TrPr">
 *   <xsd:complexContent>
 *     <xsd:extension base="CT_TrPrBase">
 *       <xsd:sequence>
 *         <xsd:element name="ins" type="CT_TrackChange" minOccurs="0"/>
 *         <xsd:element name="del" type="CT_TrackChange" minOccurs="0"/>
 *         <xsd:element name="trPrChange" type="CT_TrPrChange" minOccurs="0"/>
 *       </xsd:sequence>
 *     </xsd:extension>
 *   </xsd:complexContent>
 * </xsd:complexType>
 * <xsd:complexType name="CT_TrPrChange">
 *   <xsd:complexContent>
 *     <xsd:extension base="CT_TrackChange">
 *       <xsd:sequence>
 *         <xsd:element name="trPr" type="CT_TrPrBase" minOccurs="1"/>
 *       </xsd:sequence>
 *     </xsd:extension>
 *   </xsd:complexContent>
 * </xsd:complexType>
 * ```
 *
 * @module
 */
import type { PositiveUniversalMeasure } from "@office-open/core";
import type { ChangedAttributesProperties } from "@shared/track-revision/track-revision";

import type { AlignmentType } from "../../paragraph";
import type { TableCellSpacingProperties } from "../table-cell-spacing";
import type { TableWidthProperties } from "../table-width";
import type { HeightRule } from "./table-row-height";

/**
 * Options for CT_Cnf — conditional formatting style.
 *
 * Each boolean attribute mirrors one of the 12 positions in the conditional-
 * format bit string; val is the canonical 12-char ST_Cnf pattern ([01]*).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Cnf">
 *   <xsd:attribute name="val" type="ST_Cnf"/>
 *   <xsd:attribute name="firstRow" type="s:ST_OnOff"/>
 *   <xsd:attribute name="lastRow" type="s:ST_OnOff"/>
 *   <xsd:attribute name="firstColumn" type="s:ST_OnOff"/>
 *   <xsd:attribute name="lastColumn" type="s:ST_OnOff"/>
 *   <xsd:attribute name="oddVBand" type="s:ST_OnOff"/>
 *   <xsd:attribute name="evenVBand" type="s:ST_OnOff"/>
 *   <xsd:attribute name="oddHBand" type="s:ST_OnOff"/>
 *   <xsd:attribute name="evenHBand" type="s:ST_OnOff"/>
 *   <xsd:attribute name="firstRowFirstColumn" type="s:ST_OnOff"/>
 *   <xsd:attribute name="firstRowLastColumn" type="s:ST_OnOff"/>
 *   <xsd:attribute name="lastRowFirstColumn" type="s:ST_OnOff"/>
 *   <xsd:attribute name="lastRowLastColumn" type="s:ST_OnOff"/>
 * </xsd:complexType>
 * ```
 */
export interface CnfStyleOptions {
  /** Conditional-format bit string (ST_Cnf: 12-char [01]*) */
  val?: string;
  firstRow?: boolean;
  lastRow?: boolean;
  firstColumn?: boolean;
  lastColumn?: boolean;
  oddVBand?: boolean;
  evenVBand?: boolean;
  oddHBand?: boolean;
  evenHBand?: boolean;
  firstRowFirstColumn?: boolean;
  firstRowLastColumn?: boolean;
  lastRowFirstColumn?: boolean;
  lastRowLastColumn?: boolean;
}

export interface TableRowPropertiesOptionsBase {
  /** Conditional formatting style (cnfStyle) */
  cnfStyle?: CnfStyleOptions;
  /** Whether the row can be split across pages (cantSplit) */
  cantSplit?: boolean;
  /** Whether the row should be repeated as a header row on each page (tblHeader) */
  tableHeader?: boolean;
  /** Row height configuration (trHeight) */
  height?: {
    /** Height value in twips or as a PositiveUniversalMeasure */
    value: number | PositiveUniversalMeasure;
    /** Height rule determining how the height value is applied (ST_HeightRule) */
    rule?: (typeof HeightRule)[keyof typeof HeightRule];
  };
  /** Spacing between cells in the row (tblCellSpacing) */
  cellSpacing?: TableCellSpacingProperties;
  /** div ID for HTML compatibility (divId) */
  divId?: number;
  /** Number of grid columns before the first cell (gridBefore) */
  gridBefore?: number;
  /** Number of grid columns after the last cell (gridAfter) */
  gridAfter?: number;
  /** Preferred width before the row (wBefore) */
  widthBefore?: TableWidthProperties;
  /** Preferred width after the row (wAfter) */
  widthAfter?: TableWidthProperties;
  /** Row alignment (jc) */
  rowAlignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
  /** Whether the row is hidden (hidden) */
  hidden?: boolean;
}

/**
 * Options for configuring table row properties.
 *
 * @see {@link TableRowProperties}
 */
export type TableRowPropertiesOptions = TableRowPropertiesOptionsBase & {
  insertion?: ChangedAttributesProperties;
  deletion?: ChangedAttributesProperties;
  revision?: TableRowPropertiesChangeOptions;
  includeIfEmpty?: boolean;
};

export type TableRowPropertiesChangeOptions = TableRowPropertiesOptionsBase &
  ChangedAttributesProperties;
