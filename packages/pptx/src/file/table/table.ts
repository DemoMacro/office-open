import { XmlComponent } from "@file/xml-components";

import type { ICellBorderOptions } from "./table-cell-properties";
import { TableGrid } from "./table-grid";
import { TableProperties } from "./table-properties";
import { TableRow, type ITableRowOptions } from "./table-row";

export interface ITableOptions {
    readonly rows: readonly ITableRowOptions[];
    readonly columnWidths?: readonly number[];
    readonly firstRow?: boolean;
    readonly lastRow?: boolean;
    readonly bandRow?: boolean;
    readonly firstCol?: boolean;
    readonly lastCol?: boolean;
    readonly bandCol?: boolean;
    readonly borders?: {
        readonly top?: ICellBorderOptions;
        readonly bottom?: ICellBorderOptions;
        readonly left?: ICellBorderOptions;
        readonly right?: ICellBorderOptions;
        readonly insideHorizontal?: ICellBorderOptions;
        readonly insideVertical?: ICellBorderOptions;
    };
}

/**
 * a:tbl — DrawingML table element.
 */
export class Table extends XmlComponent {
    public constructor(options: ITableOptions) {
        super("a:tbl");

        this.root.push(
            new TableProperties({
                firstRow: options.firstRow,
                lastRow: options.lastRow,
                bandRow: options.bandRow,
                firstCol: options.firstCol,
                lastCol: options.lastCol,
                bandCol: options.bandCol,
            }),
        );

        // Column grid
        if (options.columnWidths && options.columnWidths.length > 0) {
            this.root.push(new TableGrid([...options.columnWidths]));
        } else {
            // Default: equal widths based on first row cell count
            const colCount = options.rows[0]?.cells.length ?? 1;
            this.root.push(new TableGrid(Array.from({ length: colCount }, () => 0)));
        }

        // Rows
        for (const row of options.rows) {
            this.root.push(new TableRow(row));
        }
    }
}
