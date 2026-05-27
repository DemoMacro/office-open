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
import {
  DeletedTableRow,
  InsertedTableRow,
  buildDeletedTableRowObj,
  buildInsertedTableRowObj,
} from "@file/track-revision";
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
} from "@file/xml-components";
import type { IXmlableObject } from "@file/xml-components";
import type { PositiveUniversalMeasure } from "@util/values";
import { measurementOrPercentValue, twipsMeasureValue } from "@util/values";

import { createAlignment } from "../../paragraph";
import type { AlignmentType } from "../../paragraph";
import { createTableCellSpacing } from "../table-cell-spacing";
import type { TableCellSpacingProperties } from "../table-cell-spacing";
import { buildTableWidthObj, createTableWidthElement } from "../table-width";
import type { TableWidthProperties } from "../table-width";
import { createTableRowHeight } from "./table-row-height";
import type { HeightRule } from "./table-row-height";

/**
 * Options for CT_Cnf — conditional formatting style.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Cnf">
 *   <xsd:attribute name="val" type="s:ST_String" use="required"/>
 *   <xsd:attribute name="changed" type="s:ST_OnOff"/>
 * </xsd:complexType>
 * ```
 */
export interface CnfStyleOptions {
  /** Conditional format string (required) */
  readonly val: string;
  /** Whether the property was changed */
  readonly changed?: boolean;
}

export interface TableRowPropertiesOptionsBase {
  /** Conditional formatting style (cnfStyle) */
  readonly cnfStyle?: CnfStyleOptions;
  /** Whether the row can be split across pages (cantSplit) */
  readonly cantSplit?: boolean;
  /** Whether the row should be repeated as a header row on each page (tblHeader) */
  readonly tableHeader?: boolean;
  /** Row height configuration (trHeight) */
  readonly height?: {
    /** Height value in twips or as a PositiveUniversalMeasure */
    readonly value: number | PositiveUniversalMeasure;
    /** Height rule determining how the height value is applied */
    readonly rule: (typeof HeightRule)[keyof typeof HeightRule];
  };
  /** Spacing between cells in the row (tblCellSpacing) */
  readonly cellSpacing?: TableCellSpacingProperties;
  /** div ID for HTML compatibility (divId) */
  readonly divId?: number;
  /** Number of grid columns before the first cell (gridBefore) */
  readonly gridBefore?: number;
  /** Number of grid columns after the last cell (gridAfter) */
  readonly gridAfter?: number;
  /** Preferred width before the row (wBefore) */
  readonly widthBefore?: TableWidthProperties;
  /** Preferred width after the row (wAfter) */
  readonly widthAfter?: TableWidthProperties;
  /** Row alignment (jc) */
  readonly rowAlignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
  /** Whether the row is hidden (hidden) */
  readonly hidden?: boolean;
}

/**
 * Options for configuring table row properties.
 *
 * @see {@link TableRowProperties}
 */
export type ITableRowPropertiesOptions = TableRowPropertiesOptionsBase & {
  readonly insertion?: ChangedAttributesProperties;
  readonly deletion?: ChangedAttributesProperties;
  readonly revision?: ITableRowPropertiesChangeOptions;
  readonly includeIfEmpty?: boolean;
};

export type ITableRowPropertiesChangeOptions = TableRowPropertiesOptionsBase &
  ChangedAttributesProperties;

/**
 * Build table row properties change (w:trPrChange) as IXmlableObject without allocating XmlComponent tree.
 */
export function buildTableRowPropertiesChangeObj(
  options: ITableRowPropertiesChangeOptions,
): IXmlableObject {
  const innerPr = buildTableRowProperties({ ...options, includeIfEmpty: true })!;
  return {
    "w:trPrChange": [
      { _attr: { "w:author": options.author, "w:date": options.date, "w:id": options.id } },
      innerPr,
    ],
  };
}

/**
 * Build table row properties (w:trPr) as IXmlableObject without allocating XmlComponent tree.
 */
export function buildTableRowProperties(
  options: ITableRowPropertiesOptions,
): IXmlableObject | undefined {
  const children: IXmlableObject[] = [];

  if (options.cnfStyle !== undefined) {
    const attrs: Record<string, string | boolean> = { "w:val": options.cnfStyle.val };
    if (options.cnfStyle.changed !== undefined) {
      attrs["w:changed"] = options.cnfStyle.changed;
    }
    children.push({ "w:cnfStyle": { _attr: attrs } });
  }

  if (options.divId !== undefined) {
    children.push(numberValObj("w:divId", options.divId));
  }

  if (options.gridBefore !== undefined) {
    children.push(numberValObj("w:gridBefore", options.gridBefore));
  }

  if (options.gridAfter !== undefined) {
    children.push(numberValObj("w:gridAfter", options.gridAfter));
  }

  if (options.widthBefore) {
    children.push(buildTableWidthObj("w:wBefore", options.widthBefore));
  }

  if (options.widthAfter) {
    children.push(buildTableWidthObj("w:wAfter", options.widthAfter));
  }

  if (options.cantSplit !== undefined) {
    children.push(onOffObj("w:cantSplit", options.cantSplit));
  }

  if (options.tableHeader !== undefined) {
    children.push(onOffObj("w:tblHeader", options.tableHeader));
  }

  if (options.height) {
    children.push(
      attrObj("w:trHeight", {
        "w:val": twipsMeasureValue(options.height.value),
        "w:hRule": options.height.rule,
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

  if (options.rowAlignment) {
    children.push(stringEnumValObj("w:jc", options.rowAlignment));
  }

  if (options.hidden !== undefined) {
    children.push(onOffObj("w:hidden", options.hidden));
  }

  if (options.insertion) {
    children.push(buildInsertedTableRowObj(options.insertion));
  }

  if (options.deletion) {
    children.push(buildDeletedTableRowObj(options.deletion));
  }

  if (options.revision) {
    children.push(buildTableRowPropertiesChangeObj(options.revision));
  }

  if (options.includeIfEmpty || children.length > 0) {
    return { "w:trPr": children };
  }

  return undefined;
}

/**
 * Represents table row properties (trPr) in a WordprocessingML document.
 *
 * The trPr element specifies properties for a table row including height,
 * whether it can split across pages, and whether it's a header row.
 *
 * Reference: http://officeopenxml.com/WPtableRowProperties.php
 *
 * @example
 * ```typescript
 * new TableRowProperties({
 *   cantSplit: true,
 *   tableHeader: true,
 *   height: {
 *     value: 1000,
 *     rule: HeightRule.EXACT,
 *   },
 * });
 * ```
 */
export class TableRowProperties extends IgnoreIfEmptyXmlComponent {
  public constructor(options: ITableRowPropertiesOptions) {
    super("w:trPr", options.includeIfEmpty);

    if (options.cnfStyle !== undefined) {
      const attrs: Record<string, { readonly key: string; readonly value: string | boolean }> = {
        val: { key: "w:val", value: options.cnfStyle.val },
      };
      if (options.cnfStyle.changed !== undefined) {
        attrs.changed = { key: "w:changed", value: options.cnfStyle.changed };
      }
      this.root.push(new BuilderElement({ name: "w:cnfStyle", attributes: attrs }));
    }

    if (options.divId !== undefined) {
      this.root.push(
        new BuilderElement<{ readonly val: number }>({
          name: "w:divId",
          attributes: { val: { key: "w:val", value: options.divId } },
        }),
      );
    }

    if (options.gridBefore !== undefined) {
      this.root.push(
        new BuilderElement<{ readonly val: number }>({
          name: "w:gridBefore",
          attributes: { val: { key: "w:val", value: options.gridBefore } },
        }),
      );
    }

    if (options.gridAfter !== undefined) {
      this.root.push(
        new BuilderElement<{ readonly val: number }>({
          name: "w:gridAfter",
          attributes: { val: { key: "w:val", value: options.gridAfter } },
        }),
      );
    }

    if (options.widthBefore) {
      this.root.push(createTableWidthElement("w:wBefore", options.widthBefore));
    }

    if (options.widthAfter) {
      this.root.push(createTableWidthElement("w:wAfter", options.widthAfter));
    }

    if (options.cantSplit !== undefined) {
      this.root.push(onOffObj("w:cantSplit", options.cantSplit));
    }

    if (options.tableHeader !== undefined) {
      this.root.push(onOffObj("w:tblHeader", options.tableHeader));
    }

    if (options.height) {
      this.root.push(createTableRowHeight(options.height.value, options.height.rule));
    }

    if (options.cellSpacing) {
      this.root.push(createTableCellSpacing(options.cellSpacing));
    }

    if (options.rowAlignment) {
      this.root.push(createAlignment(options.rowAlignment));
    }

    if (options.hidden !== undefined) {
      this.root.push(onOffObj("w:hidden", options.hidden));
    }

    if (options.insertion) {
      this.root.push(new InsertedTableRow(options.insertion));
    }

    if (options.deletion) {
      this.root.push(new DeletedTableRow(options.deletion));
    }

    if (options.revision) {
      this.root.push(new TableRowPropertiesChange(options.revision));
    }
  }
}

export class TableRowPropertiesChange extends XmlComponent {
  public constructor(options: ITableRowPropertiesChangeOptions) {
    super("w:trPrChange");
    this.root.push(
      new ChangeAttributes({
        author: options.author,
        date: options.date,
        id: options.id,
      }),
    );
    // TrPr is required (minOccurs="1") even if empty
    this.root.push(new TableRowProperties({ ...options, includeIfEmpty: true }));
  }
}
