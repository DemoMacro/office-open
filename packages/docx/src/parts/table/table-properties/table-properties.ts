import type { ShadingAttributesProperties } from "@shared/shading";
/**
 * Table properties module for WordprocessingML documents.
 *
 * This module provides table-level properties including width, borders,
 * layout, alignment, and margins.
 *
 * Reference: http://officeopenxml.com/WPtableProperties.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_TblPrBase">
 *   <xsd:sequence>
 *     <xsd:element name="tblStyle" type="CT_String" minOccurs="0"/>
 *     <xsd:element name="tblpPr" type="CT_TblPPr" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="tblOverlap" type="CT_TblOverlap" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="bidiVisual" type="CT_OnOff" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="tblStyleRowBandSize" type="CT_DecimalNumber" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="tblStyleColBandSize" type="CT_DecimalNumber" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="tblW" type="CT_TblWidth" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="jc" type="CT_JcTable" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="tblCellSpacing" type="CT_TblWidth" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="tblInd" type="CT_TblWidth" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="tblBorders" type="CT_TblBorders" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="shd" type="CT_Shd" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="tblLayout" type="CT_TblLayoutType" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="tblCellMar" type="CT_TblCellMar" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="tblLook" type="CT_TblLook" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="tblCaption" type="CT_String" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="tblDescription" type="CT_String" minOccurs="0" maxOccurs="1"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 *
 * <xsd:complexType name="CT_TblPrChange">
 *   <xsd:complexContent>
 *     <xsd:extension base="CT_TrackChange">
 *       <xsd:sequence>
 *         <xsd:element name="tblPr" type="CT_TblPrBase"/>
 *       </xsd:sequence>
 *     </xsd:extension>
 *   </xsd:complexContent>
 * </xsd:complexType>
 * ```
 *
 * @module
 */
import type { ChangedAttributesProperties } from "@shared/track-revision/track-revision";

import type { AlignmentType } from "../../paragraph";
import type { TableCellSpacingProperties } from "../table-cell-spacing";
import type { TableWidthProperties } from "../table-width";
import type { TableBordersOptions } from "./table-borders";
import type { TableCellMarginOptions } from "./table-cell-margin";
import type { TableFloatOptions } from "./table-float-properties";
import type { TableLayoutType } from "./table-layout";
import type { TableLookOptions } from "./table-look";

export interface TablePropertiesOptionsBase {
  width?: TableWidthProperties;
  indent?: TableWidthProperties;
  layout?: (typeof TableLayoutType)[keyof typeof TableLayoutType];
  borders?: TableBordersOptions;
  float?: TableFloatOptions;
  shading?: ShadingAttributesProperties;
  style?: string;
  alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
  cellMargin?: TableCellMarginOptions;
  visuallyRightToLeft?: boolean;
  tableLook?: TableLookOptions;
  cellSpacing?: TableCellSpacingProperties;
  /** Number of rows in each band for table style (tblStyleRowBandSize) */
  styleRowBandSize?: number;
  /** Number of columns in each band for table style (tblStyleColBandSize) */
  styleColBandSize?: number;
  /** Table caption for accessibility (tblCaption) */
  caption?: string;
  /** Table description for accessibility (tblDescription) */
  description?: string;
}

export type TablePropertiesChangeOptions = TablePropertiesOptions & ChangedAttributesProperties;

/**
 * Options for configuring table properties.
 *
 * @see {@link TableProperties}
 */
export type TablePropertiesOptions = {
  revision?: TablePropertiesChangeOptions;
  includeIfEmpty?: boolean;
} & TablePropertiesOptionsBase;
