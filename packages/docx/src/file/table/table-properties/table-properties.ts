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
import { ChangeAttributes } from "@file/track-revision/track-revision";
import type { ChangedAttributesProperties } from "@file/track-revision/track-revision";
import {
  BuilderElement,
  IgnoreIfEmptyXmlComponent,
  XmlComponent,
  attrObj,
  numberValObj,
  onOffObj,
  stringEnumValObj,
  stringValObj,
} from "@file/xml-components";
import type { IXmlableObject } from "@file/xml-components";
import { measurementOrPercentValue } from "@util/values";

import { createAlignment } from "../../paragraph";
import type { AlignmentType } from "../../paragraph";
import { buildShadingObj, createShading } from "../../shading";
import type { ShadingAttributesProperties } from "../../shading";
import { createTableCellSpacing } from "../table-cell-spacing";
import type { TableCellSpacingProperties } from "../table-cell-spacing";
import { buildTableWidthObj, createTableWidthElement } from "../table-width";
import type { TableWidthProperties } from "../table-width";
import { TableBorders, buildTableBorders } from "./table-borders";
import type { TableBordersOptions } from "./table-borders";
import { buildTableCellMarginObj, createTableCellMargin } from "./table-cell-margin";
import type { TableCellMarginOptions } from "./table-cell-margin";
import {
  buildTableFloatPropertiesObj,
  createTableFloatProperties,
  createTableOverlap,
} from "./table-float-properties";
import type { TableFloatOptions } from "./table-float-properties";
import { createTableLayout } from "./table-layout";
import type { TableLayoutType } from "./table-layout";
import { createTableLook } from "./table-look";
import type { TableLookOptions } from "./table-look";

export interface TablePropertiesOptionsBase {
  readonly width?: TableWidthProperties;
  readonly indent?: TableWidthProperties;
  readonly layout?: (typeof TableLayoutType)[keyof typeof TableLayoutType];
  readonly borders?: TableBordersOptions;
  readonly float?: TableFloatOptions;
  readonly shading?: ShadingAttributesProperties;
  readonly style?: string;
  readonly alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
  readonly cellMargin?: TableCellMarginOptions;
  readonly visuallyRightToLeft?: boolean;
  readonly tableLook?: TableLookOptions;
  readonly cellSpacing?: TableCellSpacingProperties;
  /** Number of rows in each band for table style (tblStyleRowBandSize) */
  readonly styleRowBandSize?: number;
  /** Number of columns in each band for table style (tblStyleColBandSize) */
  readonly styleColBandSize?: number;
  /** Table caption for accessibility (tblCaption) */
  readonly caption?: string;
  /** Table description for accessibility (tblDescription) */
  readonly description?: string;
}

export type ITablePropertiesChangeOptions = ITablePropertiesOptions & ChangedAttributesProperties;

/**
 * Options for configuring table properties.
 *
 * @see {@link TableProperties}
 */
export type ITablePropertiesOptions = {
  readonly revision?: ITablePropertiesChangeOptions;
  readonly includeIfEmpty?: boolean;
} & TablePropertiesOptionsBase;

/**
 * Build table properties change (w:tblPrChange) as IXmlableObject without allocating XmlComponent tree.
 */
function buildTablePropertiesChangeObj(options: ITablePropertiesChangeOptions): IXmlableObject {
  const innerPr = buildTableProperties({ ...options, includeIfEmpty: true })!;
  return {
    "w:tblPrChange": [
      { _attr: { "w:author": options.author, "w:date": options.date, "w:id": options.id } },
      innerPr,
    ],
  };
}

/**
 * Build table properties (w:tblPr) as IXmlableObject without allocating XmlComponent tree.
 */
export function buildTableProperties(options: ITablePropertiesOptions): IXmlableObject | undefined {
  const children: IXmlableObject[] = [];

  if (options.style) {
    children.push(stringValObj("w:tblStyle", options.style));
  }

  if (options.float) {
    children.push(buildTableFloatPropertiesObj(options.float));
    if (options.float.overlap) {
      children.push(stringEnumValObj("w:tblOverlap", options.float.overlap));
    }
  }

  if (options.visuallyRightToLeft !== undefined) {
    children.push(onOffObj("w:bidiVisual", options.visuallyRightToLeft));
  }

  if (options.styleRowBandSize !== undefined) {
    children.push(numberValObj("w:tblStyleRowBandSize", options.styleRowBandSize));
  }

  if (options.styleColBandSize !== undefined) {
    children.push(numberValObj("w:tblStyleColBandSize", options.styleColBandSize));
  }

  if (options.width) {
    children.push(buildTableWidthObj("w:tblW", options.width));
  }

  if (options.alignment) {
    children.push(stringEnumValObj("w:jc", options.alignment));
  }

  if (options.indent) {
    children.push(buildTableWidthObj("w:tblInd", options.indent));
  }

  if (options.borders) {
    children.push(buildTableBorders(options.borders));
  }

  if (options.shading) {
    children.push(buildShadingObj(options.shading));
  }

  if (options.layout) {
    children.push(stringEnumValObj("w:tblLayout", options.layout));
  }

  if (options.cellMargin) {
    const cellMargin = buildTableCellMarginObj(options.cellMargin);
    if (cellMargin) {
      children.push(cellMargin);
    }
  }

  if (options.tableLook) {
    children.push(
      attrObj("w:tblLook", {
        "w:firstRow": options.tableLook.firstRow,
        "w:lastRow": options.tableLook.lastRow,
        "w:firstColumn": options.tableLook.firstColumn,
        "w:lastColumn": options.tableLook.lastColumn,
        "w:noHBand": options.tableLook.noHBand,
        "w:noVBand": options.tableLook.noVBand,
      }),
    );
  }

  if (options.cellSpacing) {
    children.push(
      attrObj("w:tblCellSpacing", {
        "w:w": measurementOrPercentValue(options.cellSpacing.value),
        "w:type": options.cellSpacing.type,
      }),
    );
  }

  if (options.caption !== undefined) {
    children.push(stringValObj("w:tblCaption", options.caption));
  }

  if (options.description !== undefined) {
    children.push(stringValObj("w:tblDescription", options.description));
  }

  if (options.revision) {
    children.push(buildTablePropertiesChangeObj(options.revision));
  }

  if (options.includeIfEmpty || children.length > 0) {
    return { "w:tblPr": children };
  }

  return undefined;
}

/**
 * Represents table properties (tblPr) in a WordprocessingML document.
 *
 * The tblPr element specifies the properties for a table including width,
 * alignment, borders, margins, and layout.
 *
 * Reference: http://officeopenxml.com/WPtableProperties.php
 */
export class TableProperties extends IgnoreIfEmptyXmlComponent {
  public constructor(options: ITablePropertiesOptions) {
    super("w:tblPr", options.includeIfEmpty);

    if (options.style) {
      this.root.push(stringValObj("w:tblStyle", options.style));
    }

    if (options.float) {
      this.root.push(createTableFloatProperties(options.float));
      if (options.float.overlap) {
        this.root.push(createTableOverlap(options.float.overlap));
      }
    }

    if (options.visuallyRightToLeft !== undefined) {
      this.root.push(onOffObj("w:bidiVisual", options.visuallyRightToLeft));
    }

    if (options.styleRowBandSize !== undefined) {
      this.root.push(
        new BuilderElement<{ readonly val: number }>({
          name: "w:tblStyleRowBandSize",
          attributes: { val: { key: "w:val", value: options.styleRowBandSize } },
        }),
      );
    }

    if (options.styleColBandSize !== undefined) {
      this.root.push(
        new BuilderElement<{ readonly val: number }>({
          name: "w:tblStyleColBandSize",
          attributes: { val: { key: "w:val", value: options.styleColBandSize } },
        }),
      );
    }

    if (options.width) {
      this.root.push(createTableWidthElement("w:tblW", options.width));
    }

    if (options.alignment) {
      this.root.push(createAlignment(options.alignment));
    }

    if (options.indent) {
      this.root.push(createTableWidthElement("w:tblInd", options.indent));
    }

    if (options.borders) {
      this.root.push(new TableBorders(options.borders));
    }

    if (options.shading) {
      this.root.push(createShading(options.shading));
    }

    if (options.layout) {
      this.root.push(createTableLayout(options.layout));
    }

    if (options.cellMargin) {
      const cellMargin = createTableCellMargin(options.cellMargin);
      if (cellMargin) {
        this.root.push(cellMargin);
      }
    }

    if (options.tableLook) {
      this.root.push(createTableLook(options.tableLook));
    }

    if (options.cellSpacing) {
      this.root.push(createTableCellSpacing(options.cellSpacing));
    }

    if (options.caption !== undefined) {
      this.root.push(stringValObj("w:tblCaption", options.caption));
    }

    if (options.description !== undefined) {
      this.root.push(stringValObj("w:tblDescription", options.description));
    }

    if (options.revision) {
      this.root.push(new TablePropertiesChange(options.revision));
    }
  }
}

class TablePropertiesChange extends XmlComponent {
  public constructor(options: ITablePropertiesChangeOptions) {
    super("w:tblPrChange");
    this.root.push(
      new ChangeAttributes({
        author: options.author,
        date: options.date,
        id: options.id,
      }),
    );
    // TblPr is required even if empty (minOccurs="0" is missing)
    this.root.push(new TableProperties({ ...options, includeIfEmpty: true }));
  }
}
