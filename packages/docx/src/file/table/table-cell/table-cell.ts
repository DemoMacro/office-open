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
import type { Context, IXmlableObject } from "@file/xml-components";
import { xml } from "@office-open/xml";

import type { StructuredDocumentTagCell } from "../../sdt";
import type { SectionChild } from "../../section-child";
import { buildTableCellProperties } from "./table-cell-properties";
import type { ITableCellPropertiesOptions } from "./table-cell-properties";

// Lazy import to avoid circular dependency: coerce.ts → Table → TableRow → TableCell → coerce.ts
type CoerceFn = (child: SectionChild) => FileChild;
let _coerce: CoerceFn | undefined;

function lazyCoerce(child: SectionChild): FileChild {
  if (!_coerce) {
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    _coerce = require("../../coerce").coerceSectionChild;
  }
  return _coerce!(child);
}

/**
 * Options for creating a TableCell element.
 *
 * @see {@link TableCell}
 */
export type TableCellOptions = {
  /** Array of Paragraph, Table, or plain objects that make up the cell content */
  readonly children: readonly (SectionChild | StructuredDocumentTagCell)[];
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
  public constructor(public readonly options: TableCellOptions) {
    super("w:tc");
  }

  public prepForXml(context: Context): IXmlableObject | undefined {
    const children: IXmlableObject[] = [];

    const tPrObj = buildTableCellProperties(this.options);
    if (tPrObj) children.push(tPrObj);

    const coerced = this.options.children.map((c) =>
      c instanceof BaseXmlComponent ? c : lazyCoerce(c),
    );

    for (const child of coerced) {
      const obj = child.prepForXml(context);
      if (obj) children.push(obj);
    }

    // Cells must end with a paragraph
    const last = coerced[coerced.length - 1];
    if (!(last instanceof Paragraph)) {
      children.push(new Paragraph({}).prepForXml(context)!);
    }

    return { "w:tc": children };
  }

  public override toXml(context: Context): string {
    const parts: string[] = [];

    const tPrObj = buildTableCellProperties(this.options);
    if (tPrObj) parts.push(xml(tPrObj));

    const coerced = this.options.children.map((c) =>
      c instanceof BaseXmlComponent ? c : lazyCoerce(c),
    );

    for (const child of coerced) {
      parts.push(child.toXml(context));
    }

    // Cells must end with a paragraph
    const last = coerced[coerced.length - 1];
    if (!(last instanceof Paragraph)) {
      parts.push(new Paragraph({}).toXml(context));
    }

    return `<w:tc>${parts.join("")}</w:tc>`;
  }
}
