/**
 * Table module for WordprocessingML documents.
 *
 * Reference: http://officeopenxml.com/WPtableGrid.php
 *
 * @module
 */
import { FileChild } from "@file/file-child";

import type { AlignmentType } from "../paragraph";
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
import type { TableRow } from "./table-row";
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
    readonly rows: readonly TableRow[];
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
export class Table extends FileChild {
    public constructor({
        rows,
        width,
        columnWidths = Array<number>(Math.max(...rows.map((row) => row.CellCount))).fill(100),
        columnWidthsRevision,
        margins,
        indent,
        float,
        layout,
        style,
        borders,
        alignment,
        visuallyRightToLeft,
        tableLook,
        cellSpacing,
        styleRowBandSize,
        styleColBandSize,
        caption,
        description,
        revision,
    }: ITableOptions) {
        super("w:tbl");

        this.root.push(
            new TableProperties({
                alignment,
                borders: borders ?? {},
                caption,
                cellMargin: margins,
                cellSpacing,
                description,
                float,
                indent,
                layout,
                revision,
                style,
                styleColBandSize,
                styleRowBandSize,
                tableLook,
                visuallyRightToLeft,
                width: width ?? { size: 100 },
            }),
        );

        this.root.push(new TableGrid(columnWidths, columnWidthsRevision));

        for (const row of rows) {
            this.root.push(row);
        }

        rows.forEach((row, rowIndex) => {
            if (rowIndex === rows.length - 1) {
                // Don't process the end row
                return;
            }
            let columnIndex = 0;
            row.cells.forEach((cell) => {
                // Row Span has to be added in this method and not the constructor because it needs to know information about the column which happens after Table Cell construction
                // Row Span of 1 will crash word as it will add RESTART and not a corresponding CONTINUE
                if (cell.options.rowSpan && cell.options.rowSpan > 1) {
                    const continueCell = new TableCell({
                        // The inserted CONTINUE cell has rowSpan, and will be handled when process the next row
                        borders: cell.options.borders,
                        children: [],
                        columnSpan: cell.options.columnSpan,
                        rowSpan: cell.options.rowSpan - 1,
                        verticalMerge: VerticalMergeType.CONTINUE,
                    });
                    rows[rowIndex + 1].addCellToColumnIndex(continueCell, columnIndex);
                }
                columnIndex += cell.options.columnSpan || 1;
            });
        });
    }
}
