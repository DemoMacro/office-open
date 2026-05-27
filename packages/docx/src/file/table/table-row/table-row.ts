/**
 * Table row module for WordprocessingML documents.
 *
 * Reference: http://officeopenxml.com/WPtableRow.php
 *
 * @module
 */
import { BaseXmlComponent, EMPTY_OBJECT } from "@file/xml-components";
import type { Context, IXmlableObject } from "@file/xml-components";

import { StructuredDocumentTagCell } from "../../sdt";
import { StructuredDocumentTagRow } from "../../sdt";
import { TableCell } from "../table-cell";
import type { TableCellOptions } from "../table-cell";
import { TablePropertyExceptions } from "../table-properties/table-property-exceptions";
import type { TablePropertyExOptions } from "../table-properties/table-property-exceptions";
import { buildTableRowProperties } from "./table-row-properties";
import type { ITableRowPropertiesOptions } from "./table-row-properties";

/**
 * Options for creating a TableRow element.
 *
 * @see {@link TableRow}
 */
export type TableRowOptions = {
  /** Array of TableCell elements or plain options that make up the row */
  readonly children: readonly (
    | TableCell
    | StructuredDocumentTagCell
    | StructuredDocumentTagRow
    | TableCellOptions
  )[];
  /** Table property exceptions for this row (override table-level properties) */
  readonly propertyExceptions?: TablePropertyExOptions;
} & ITableRowPropertiesOptions;

/**
 * Represents a table row in a WordprocessingML document.
 *
 * A table row is a single row of cells within a table. Each row contains
 * one or more table cells that hold the actual content.
 *
 * Reference: http://officeopenxml.com/WPtableRow.php
 *
 * @publicApi
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Row">
 *   <xsd:sequence>
 *     <xsd:element name="tblPrEx" type="CT_TblPrEx" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="trPr" type="CT_TrPr" minOccurs="0" maxOccurs="1"/>
 *     <xsd:group ref="EG_ContentCellContent" minOccurs="0" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="rsidRPr" type="ST_LongHexNumber"/>
 *   <xsd:attribute name="rsidR" type="ST_LongHexNumber"/>
 *   <xsd:attribute name="rsidDel" type="ST_LongHexNumber"/>
 *   <xsd:attribute name="rsidTr" type="ST_LongHexNumber"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * new TableRow({
 *   children: [
 *     new TableCell({ children: [new Paragraph("Cell 1")] }),
 *     new TableCell({ children: [new Paragraph("Cell 2")] }),
 *   ],
 * });
 * ```
 */
export class TableRow extends BaseXmlComponent {
  // Extra cells inserted by Table's row-span CONTINUE logic
  private extraCells: { cell: TableCell; columnIndex: number }[] = [];

  // Coerced children: plain TableCellOptions are converted to TableCell instances
  private readonly coercedChildren: readonly (
    | TableCell
    | StructuredDocumentTagCell
    | StructuredDocumentTagRow
  )[];

  public constructor(private readonly options: TableRowOptions) {
    super("w:tr");
    this.coercedChildren = options.children.map((child) =>
      child instanceof TableCell ||
      child instanceof StructuredDocumentTagCell ||
      child instanceof StructuredDocumentTagRow
        ? child
        : new TableCell(child),
    );
  }

  public get CellCount(): number {
    return this.coercedChildren.length;
  }

  public get cells(): readonly TableCell[] {
    return this.coercedChildren.filter((child): child is TableCell => child instanceof TableCell);
  }

  /** @internal Used by Table to insert CONTINUE cells for vertical merge. */
  public addCellToColumnIndex(cell: TableCell, columnIndex: number): void {
    this.extraCells.push({ cell, columnIndex });
  }

  public rootIndexToColumnIndex(rootIndex: number): number {
    // rootIndex is 0-based in XmlComponent root (0 = trPr, 1+ = cells)
    const cellIndex = rootIndex - 1;
    if (cellIndex < 0 || cellIndex >= this.coercedChildren.length) {
      throw new Error(`cell 'rootIndex' should between 1 to ${this.coercedChildren.length}`);
    }
    let colIdx = 0;
    for (let i = 0; i < cellIndex; i++) {
      const child = this.coercedChildren[i];
      colIdx += child instanceof TableCell ? child.options.columnSpan || 1 : 1;
    }
    return colIdx;
  }

  public columnIndexToRootIndex(columnIndex: number, allowEndNewCell: boolean = false): number {
    if (columnIndex < 0) {
      throw new Error(`cell 'columnIndex' should not less than zero`);
    }
    let colIdx = 0;
    let idx = 0;
    while (colIdx <= columnIndex) {
      if (idx >= this.coercedChildren.length) {
        if (allowEndNewCell) {
          return this.coercedChildren.length + 1;
        } else {
          throw new Error(`cell 'columnIndex' should not great than ${colIdx - 1}`);
        }
      }
      const child = this.coercedChildren[idx];
      idx += 1;
      colIdx += child instanceof TableCell ? child.options.columnSpan || 1 : 1;
    }
    return idx;
  }

  public prepForXml(context: Context): IXmlableObject | undefined {
    const children: IXmlableObject[] = [];

    if (this.options.propertyExceptions) {
      const obj = new TablePropertyExceptions(this.options.propertyExceptions).prepForXml(context);
      if (obj) children.push(obj);
    }

    const trPrObj = buildTableRowProperties(this.options);
    if (trPrObj) children.push(trPrObj);

    const prefixCount = children.length;

    for (const child of this.coercedChildren) {
      const obj = child.prepForXml(context);
      if (obj) children.push(obj);
    }

    // Insert extra CONTINUE cells at their designated column positions
    if (this.extraCells.length > 0) {
      for (const { cell, columnIndex } of this.extraCells) {
        const insertIdx = this.findInsertIndex(columnIndex, prefixCount);
        const obj = cell.prepForXml(context);
        if (obj) children.splice(insertIdx, 0, obj);
      }
    }

    return { "w:tr": children.length ? children : EMPTY_OBJECT };
  }

  private findInsertIndex(columnIndex: number, prefixCount: number): number {
    let colIdx = 0;
    for (let i = 0; i < this.coercedChildren.length; i++) {
      const child = this.coercedChildren[i];
      colIdx += child instanceof TableCell ? child.options.columnSpan || 1 : 1;
      if (colIdx > columnIndex) {
        return i + prefixCount;
      }
    }
    return this.coercedChildren.length + prefixCount;
  }
}
