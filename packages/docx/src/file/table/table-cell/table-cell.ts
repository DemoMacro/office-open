/**
 * Table cell module for WordprocessingML documents.
 *
 * Reference: http://officeopenxml.com/WPtableCell.php
 *
 * @module
 */
import type { FileChild } from "@file/file-child";
import { Paragraph } from "@file/paragraph";
import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { xml } from "@office-open/xml";

import type { StructuredDocumentTagCell } from "../../sdt";
import type { SectionChild } from "../../section-child";
import { buildTableCellProperties } from "./table-cell-properties";
import type { ITableCellPropertiesOptions } from "./table-cell-properties";

/**
 * Options for creating a TableCell element.
 *
 * @see {@link TableCell}
 */
export type TableCellOptions = {
  /** Array of Paragraph, Table, or plain objects that make up the cell content */
  children: (SectionChild | StructuredDocumentTagCell)[];
} & ITableCellPropertiesOptions;

/**
 * Represents a table cell in a WordprocessingML document.
 *
 * A table cell is the basic unit of content within a table. Each cell can contain
 * paragraphs, nested tables, or other block-level content. Cells must always end
 * with a paragraph element.
 *
 * Reference: http://officeopenxml.com/WPtableCell.php
 *
 * @publicApi
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Tc">
 *   <xsd:sequence>
 *     <xsd:element name="tcPr" type="CT_TcPr" minOccurs="0" maxOccurs="1"/>
 *     <xsd:group ref="EG_BlockLevelElts" minOccurs="1" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="id" type="s:ST_String" use="optional"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * new TableCell({
 *   children: [new Paragraph("Cell content")],
 *   width: { size: 3000, type: WidthType.DXA },
 * });
 * ```
 */
export class TableCell extends BaseXmlComponent {
  public constructor(public options: TableCellOptions) {
    super("w:tc");
  }

  public override toXml(context: Context): string {
    const parts: string[] = [];

    const tPrObj = buildTableCellProperties(this.options);
    if (tPrObj) parts.push(xml(tPrObj));

    // Children are pre-coerced by coerceSectionChild, so all should be instances.
    // Cast is safe: plain objects are converted before reaching this point.
    const children = this.options.children as (FileChild | StructuredDocumentTagCell)[];

    for (const child of children) {
      parts.push(child.toXml(context));
    }

    // Cells must end with a paragraph
    const last = children[children.length - 1];
    if (!(last instanceof Paragraph)) {
      parts.push(new Paragraph({}).toXml(context));
    }

    return `<w:tc>${parts.join("")}</w:tc>`;
  }
}
