/**
 * Conditional formatting style module for WordprocessingML documents.
 *
 * This module provides the conditional formatting style element (w:cnfStyle)
 * which specifies conditional formatting for table cells based on their position.
 *
 * Reference: http://officeopenxml.com/WPtableCellProperties.php
 *
 * @module
 */
import { BuilderElement, attrObj } from "@file/xml-components";
import type { IXmlableObject, XmlComponent } from "@file/xml-components";

/**
 * Options for conditional formatting style.
 *
 * These options specify which conditions apply to the element based on
 * its position within a table (first/last row, first/last column, bands).
 */
export interface CnfConditionalOptions {
  /** Whether this is the first row in the table */
  firstRow?: boolean;
  /** Whether this is the last row in the table */
  lastRow?: boolean;
  /** Whether this is the first column in the table */
  firstColumn?: boolean;
  /** Whether this is the last column in the table */
  lastColumn?: boolean;
  /** Whether this is an odd vertical band */
  oddVBand?: boolean;
  /** Whether this is an even vertical band */
  evenVBand?: boolean;
  /** Whether this is an odd horizontal band */
  oddHBand?: boolean;
  /** Whether this is an even horizontal band */
  evenHBand?: boolean;
  /** Whether this is the first row and first column (top-left corner) */
  firstRowFirstColumn?: boolean;
  /** Whether this is the first row and last column (top-right corner) */
  firstRowLastColumn?: boolean;
  /** Whether this is the last row and first column (bottom-left corner) */
  lastRowFirstColumn?: boolean;
  /** Whether this is the last row and last column (bottom-right corner) */
  lastRowLastColumn?: boolean;
}

/**
 * Creates a conditional formatting style element (w:cnfStyle).
 *
 * This element specifies conditional formatting properties based on the
 * position of a cell within a table. It's used to apply different styles
 * to header rows, footer rows, banded rows/columns, and corner cells.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Cnf">
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
 *
 * @example
 * ```typescript
 * // Apply style to first row
 * createCnfStyle({ firstRow: true });
 *
 * // Apply style to odd horizontal bands
 * createCnfStyle({ oddHBand: true });
 *
 * // Apply style to top-left corner cell
 * createCnfStyle({ firstRow: true, firstColumn: true, firstRowFirstColumn: true });
 * ```
 */
export const createCnfStyle = (options: CnfConditionalOptions): XmlComponent => {
  const attributes: Record<string, { key: string; value: string }> = {};

  if (options.firstRow !== undefined) {
    attributes.firstRow = { key: "w:firstRow", value: options.firstRow ? "1" : "0" };
  }
  if (options.lastRow !== undefined) {
    attributes.lastRow = { key: "w:lastRow", value: options.lastRow ? "1" : "0" };
  }
  if (options.firstColumn !== undefined) {
    attributes.firstColumn = { key: "w:firstColumn", value: options.firstColumn ? "1" : "0" };
  }
  if (options.lastColumn !== undefined) {
    attributes.lastColumn = { key: "w:lastColumn", value: options.lastColumn ? "1" : "0" };
  }
  if (options.oddVBand !== undefined) {
    attributes.oddVBand = { key: "w:oddVBand", value: options.oddVBand ? "1" : "0" };
  }
  if (options.evenVBand !== undefined) {
    attributes.evenVBand = { key: "w:evenVBand", value: options.evenVBand ? "1" : "0" };
  }
  if (options.oddHBand !== undefined) {
    attributes.oddHBand = { key: "w:oddHBand", value: options.oddHBand ? "1" : "0" };
  }
  if (options.evenHBand !== undefined) {
    attributes.evenHBand = { key: "w:evenHBand", value: options.evenHBand ? "1" : "0" };
  }
  if (options.firstRowFirstColumn !== undefined) {
    attributes.firstRowFirstColumn = {
      key: "w:firstRowFirstColumn",
      value: options.firstRowFirstColumn ? "1" : "0",
    };
  }
  if (options.firstRowLastColumn !== undefined) {
    attributes.firstRowLastColumn = {
      key: "w:firstRowLastColumn",
      value: options.firstRowLastColumn ? "1" : "0",
    };
  }
  if (options.lastRowFirstColumn !== undefined) {
    attributes.lastRowFirstColumn = {
      key: "w:lastRowFirstColumn",
      value: options.lastRowFirstColumn ? "1" : "0",
    };
  }
  if (options.lastRowLastColumn !== undefined) {
    attributes.lastRowLastColumn = {
      key: "w:lastRowLastColumn",
      value: options.lastRowLastColumn ? "1" : "0",
    };
  }

  return new BuilderElement({
    attributes,
    name: "w:cnfStyle",
  });
};

export const buildCnfStyleObj = (options: CnfConditionalOptions): IXmlableObject =>
  attrObj("w:cnfStyle", {
    "w:firstRow": options.firstRow !== undefined ? (options.firstRow ? "1" : "0") : undefined,
    "w:lastRow": options.lastRow !== undefined ? (options.lastRow ? "1" : "0") : undefined,
    "w:firstColumn":
      options.firstColumn !== undefined ? (options.firstColumn ? "1" : "0") : undefined,
    "w:lastColumn": options.lastColumn !== undefined ? (options.lastColumn ? "1" : "0") : undefined,
    "w:oddVBand": options.oddVBand !== undefined ? (options.oddVBand ? "1" : "0") : undefined,
    "w:evenVBand": options.evenVBand !== undefined ? (options.evenVBand ? "1" : "0") : undefined,
    "w:oddHBand": options.oddHBand !== undefined ? (options.oddHBand ? "1" : "0") : undefined,
    "w:evenHBand": options.evenHBand !== undefined ? (options.evenHBand ? "1" : "0") : undefined,
    "w:firstRowFirstColumn":
      options.firstRowFirstColumn !== undefined
        ? options.firstRowFirstColumn
          ? "1"
          : "0"
        : undefined,
    "w:firstRowLastColumn":
      options.firstRowLastColumn !== undefined
        ? options.firstRowLastColumn
          ? "1"
          : "0"
        : undefined,
    "w:lastRowFirstColumn":
      options.lastRowFirstColumn !== undefined
        ? options.lastRowFirstColumn
          ? "1"
          : "0"
        : undefined,
    "w:lastRowLastColumn":
      options.lastRowLastColumn !== undefined ? (options.lastRowLastColumn ? "1" : "0") : undefined,
  });
