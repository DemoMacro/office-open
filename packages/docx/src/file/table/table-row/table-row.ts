/**
 * Table row module for WordprocessingML documents.
 *
 * Reference: http://officeopenxml.com/WPtableRow.php
 *
 * @module
 */
import { BaseXmlComponent, EMPTY_OBJECT } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";

import type { StructuredDocumentTagRow } from "../../sdt";
import { TableCell } from "../table-cell";
import { TablePropertyExceptions } from "../table-properties/table-property-exceptions";
import type { ITablePropertyExOptions } from "../table-properties/table-property-exceptions";
import { TableRowProperties } from "./table-row-properties";
import type { ITableRowPropertiesOptions } from "./table-row-properties";

/**
 * Options for creating a TableRow element.
 *
 * @see {@link TableRow}
 */
export type ITableRowOptions = {
    /** Array of TableCell elements that make up the row */
    readonly children: readonly (TableCell | StructuredDocumentTagRow)[];
    /** Table property exceptions for this row (override table-level properties) */
    readonly propertyExceptions?: ITablePropertyExOptions;
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

    public constructor(private readonly options: ITableRowOptions) {
        super("w:tr");
    }

    public get CellCount(): number {
        return this.options.children.length;
    }

    public get cells(): readonly TableCell[] {
        return this.options.children.filter(
            (child): child is TableCell => child instanceof TableCell,
        );
    }

    /** @internal Used by Table to insert CONTINUE cells for vertical merge. */
    public addCellToColumnIndex(cell: TableCell, columnIndex: number): void {
        this.extraCells.push({ cell, columnIndex });
    }

    public rootIndexToColumnIndex(rootIndex: number): number {
        // rootIndex is 0-based in XmlComponent root (0 = trPr, 1+ = cells)
        const cellIndex = rootIndex - 1;
        if (cellIndex < 0 || cellIndex >= this.options.children.length) {
            throw new Error(`cell 'rootIndex' should between 1 to ${this.options.children.length}`);
        }
        let colIdx = 0;
        for (let i = 0; i < cellIndex; i++) {
            const child = this.options.children[i];
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
            if (idx >= this.options.children.length) {
                if (allowEndNewCell) {
                    return this.options.children.length + 1;
                } else {
                    throw new Error(`cell 'columnIndex' should not great than ${colIdx - 1}`);
                }
            }
            const child = this.options.children[idx];
            idx += 1;
            colIdx += child instanceof TableCell ? child.options.columnSpan || 1 : 1;
        }
        return idx;
    }

    public prepForXml(context: IContext): IXmlableObject | undefined {
        const children: IXmlableObject[] = [];

        if (this.options.propertyExceptions) {
            const obj = new TablePropertyExceptions(this.options.propertyExceptions).prepForXml(
                context,
            );
            if (obj) children.push(obj);
        }

        const trPr = new TableRowProperties(this.options);
        const trPrObj = trPr.prepForXml(context);
        if (trPrObj) children.push(trPrObj);

        const prefixCount = children.length;

        for (const child of this.options.children) {
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
        for (let i = 0; i < this.options.children.length; i++) {
            const child = this.options.children[i];
            colIdx += child instanceof TableCell ? child.options.columnSpan || 1 : 1;
            if (colIdx > columnIndex) {
                return i + prefixCount;
            }
        }
        return this.options.children.length + prefixCount;
    }
}
