/**
 * Table module for WordprocessingML documents.
 *
 * Reference: http://officeopenxml.com/WPtableGrid.php
 *
 * @module
 */
import type { FileChild } from "@file/file-child";
import { BaseXmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";

import type { AlignmentType } from "../paragraph";
import type { StructuredDocumentTagRow } from "../sdt";
import { TableGrid } from "./grid";
import type { ITableGridChangeOptions } from "./grid";
import { TableCell, VerticalMergeType } from "./table-cell";
import type { ITableCellSpacingProperties } from "./table-cell-spacing";
import { TableProperties } from "./table-properties";
import type {
    ITableBordersOptions,
    ITableFloatOptions,
    ITablePropertiesChangeOptions,
} from "./table-properties";
import type { ITableCellMarginOptions } from "./table-properties/table-cell-margin";
import type { TableLayoutType } from "./table-properties/table-layout";
import type { ITableLookOptions } from "./table-properties/table-look";
import { TableRow } from "./table-row";
import type { ITableWidthProperties } from "./table-width";

/**
 * Options for creating a Table element.
 *
 * Note: 0-width columns don't get rendered correctly, so we need
 * to give them some value. A reasonable default would be
 * ~6in / numCols, but if we do that it becomes very hard
 * to resize the table using setWidth, unless the layout
 * algorithm is set to 'fixed'. Instead, the approach here
 * means even in 'auto' layout, setting a width on the
 * table will make it look reasonable, as the layout
 * algorithm will expand columns to fit its content.
 *
 * @see {@link Table}
 */
export interface ITableOptions {
    readonly rows: readonly (TableRow | StructuredDocumentTagRow)[];
    readonly width?: ITableWidthProperties;
    readonly columnWidths?: readonly number[];
    readonly columnWidthsRevision?: ITableGridChangeOptions;
    readonly margins?: ITableCellMarginOptions;
    readonly indent?: ITableWidthProperties;
    readonly float?: ITableFloatOptions;
    readonly layout?: (typeof TableLayoutType)[keyof typeof TableLayoutType];
    readonly style?: string;
    readonly borders?: ITableBordersOptions;
    readonly alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
    readonly visuallyRightToLeft?: boolean;
    readonly tableLook?: ITableLookOptions;
    readonly cellSpacing?: ITableCellSpacingProperties;
    readonly styleRowBandSize?: number;
    readonly styleColBandSize?: number;
    readonly caption?: string;
    readonly description?: string;
    readonly revision?: ITablePropertiesChangeOptions;
}

/**
 * Represents a table in a WordprocessingML document.
 *
 * A table is a set of paragraphs (and other block-level content) arranged in rows and columns.
 * Tables are used to organize content into a grid structure.
 *
 * Reference: http://officeopenxml.com/WPtable.php
 *
 * @publicApi
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Tbl">
 *   <xsd:sequence>
 *     <xsd:group ref="EG_RangeMarkupElements" minOccurs="0" maxOccurs="unbounded"/>
 *     <xsd:element name="tblPr" type="CT_TblPr"/>
 *     <xsd:element name="tblGrid" type="CT_TblGrid"/>
 *     <xsd:group ref="EG_ContentRowContent" minOccurs="0" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * new Table({
 *   rows: [
 *     new TableRow({
 *       children: [
 *         new TableCell({ children: [new Paragraph("Cell 1")] }),
 *         new TableCell({ children: [new Paragraph("Cell 2")] }),
 *       ],
 *     }),
 *   ],
 * });
 * ```
 */
export class Table extends BaseXmlComponent implements FileChild {
    public readonly fileChild = Symbol();

    private readonly options: ITableOptions;
    private readonly columnWidths: readonly number[];

    public constructor(options: ITableOptions) {
        super("w:tbl");
        this.options = options;
        this.columnWidths =
            options.columnWidths ??
            Array<number>(Math.max(...options.rows.map((row) => row.CellCount))).fill(100);

        // Register CONTINUE cells on subsequent rows for vertical merge
        const rows = options.rows;
        rows.forEach((row, rowIndex) => {
            if (rowIndex === rows.length - 1) return;
            if (!(row instanceof TableRow)) return;

            let columnIndex = 0;
            row.cells.forEach((cell) => {
                if (cell.options.rowSpan && cell.options.rowSpan > 1) {
                    const nextRow = rows[rowIndex + 1];
                    if (nextRow instanceof TableRow) {
                        const continueCell = new TableCell({
                            borders: cell.options.borders,
                            children: [],
                            columnSpan: cell.options.columnSpan,
                            rowSpan: cell.options.rowSpan - 1,
                            verticalMerge: VerticalMergeType.CONTINUE,
                        });
                        nextRow.addCellToColumnIndex(continueCell, columnIndex);
                    }
                }
                columnIndex += cell.options.columnSpan || 1;
            });
        });
    }

    public prepForXml(context: IContext): IXmlableObject | undefined {
        const children: IXmlableObject[] = [];

        const tblPr = new TableProperties({
            alignment: this.options.alignment,
            borders: this.options.borders ?? {},
            caption: this.options.caption,
            cellMargin: this.options.margins,
            cellSpacing: this.options.cellSpacing,
            description: this.options.description,
            float: this.options.float,
            indent: this.options.indent,
            layout: this.options.layout,
            revision: this.options.revision,
            style: this.options.style,
            styleColBandSize: this.options.styleColBandSize,
            styleRowBandSize: this.options.styleRowBandSize,
            tableLook: this.options.tableLook,
            visuallyRightToLeft: this.options.visuallyRightToLeft,
            width: this.options.width ?? { size: 100 },
        });
        const tblPrObj = tblPr.prepForXml(context);
        if (tblPrObj) children.push(tblPrObj);

        const gridObj = new TableGrid(
            this.columnWidths,
            this.options.columnWidthsRevision,
        ).prepForXml(context);
        if (gridObj) children.push(gridObj);

        for (const row of this.options.rows) {
            const obj = row.prepForXml(context);
            if (obj) children.push(obj);
        }

        return { "w:tbl": children };
    }
}
