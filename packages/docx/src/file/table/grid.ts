/**
 * Table grid module for WordprocessingML documents.
 *
 * The table grid defines the column structure of a table.
 *
 * Reference: http://officeopenxml.com/WPtableGrid.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_TblGridCol">
 *   <xsd:attribute name="w" type="s:ST_TwipsMeasure"/>
 * </xsd:complexType>
 * <xsd:complexType name="CT_TblGridBase">
 *   <xsd:sequence>
 *     <xsd:element name="gridCol" type="CT_TblGridCol" minOccurs="0" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * <xsd:complexType name="CT_TblGridChange">
 *   <xsd:complexContent>
 *     <xsd:extension base="CT_Markup">
 *       <xsd:sequence>
 *         <xsd:element name="tblGrid" type="CT_TblGridBase"/>
 *       </xsd:sequence>
 *     </xsd:extension>
 *   </xsd:complexContent>
 * </xsd:complexType>
 * ```
 *
 * @module
 */
import { BuilderElement, XmlComponent } from "@file/xml-components";
import { twipsMeasureValue } from "@util/values";
import type { PositiveUniversalMeasure } from "@util/values";

export interface TableGridChangeOptions {
  id: number;
  columnWidths: number[] | PositiveUniversalMeasure[];
}

/**
 * Creates a single column in the table grid.
 *
 * The gridCol element specifies the width of a single column.
 */
export const createGridCol = (width?: number | PositiveUniversalMeasure): XmlComponent =>
  new BuilderElement<{ width?: number | PositiveUniversalMeasure }>({
    attributes:
      width !== undefined
        ? {
            width: { key: "w:w", value: twipsMeasureValue(width) },
          }
        : undefined,
    name: "w:gridCol",
  });

/**
 * Creates the table grid for a WordprocessingML document.
 *
 * The tblGrid element defines the number and width of columns in the table.
 *
 * Reference: http://officeopenxml.com/WPtableGrid.php
 */
export class TableGrid extends XmlComponent {
  public constructor(
    widths: number[] | PositiveUniversalMeasure[],
    revision?: TableGridChangeOptions,
  ) {
    super("w:tblGrid");
    for (const width of widths) {
      this.root.push(createGridCol(width));
    }
    if (revision) {
      this.root.push(new TableGridChange(revision));
    }
  }
}

export class TableGridChange extends XmlComponent {
  public constructor(options: TableGridChangeOptions) {
    super("w:tblGridChange");
    this.root.push({ _attr: { "w:id": options.id } });
    this.root.push(new TableGrid(options.columnWidths));
  }
}
