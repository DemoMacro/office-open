import { BaseXmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";

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
 * Lazy: stores options, builds IXmlableObject in prepForXml.
 */
export class Table extends BaseXmlComponent {
    private readonly options: ITableOptions;

    public constructor(options: ITableOptions) {
        super("a:tbl");
        this.options = options;
    }

    public override prepForXml(context: IContext): IXmlableObject {
        const opts = this.options;
        const children: IXmlableObject[] = [];

        // a:tblPr
        const tblPr = new TableProperties({
            firstRow: opts.firstRow,
            lastRow: opts.lastRow,
            bandRow: opts.bandRow,
            firstCol: opts.firstCol,
            lastCol: opts.lastCol,
            bandCol: opts.bandCol,
        });
        const tblPrObj = tblPr.prepForXml(context);
        if (tblPrObj) children.push(tblPrObj);

        // a:tblGrid
        const colWidths =
            opts.columnWidths && opts.columnWidths.length > 0
                ? [...opts.columnWidths]
                : Array.from({ length: opts.rows[0]?.cells.length ?? 1 }, () => 0);
        const grid = new TableGrid(colWidths);
        const gridObj = grid.prepForXml(context);
        if (gridObj) children.push(gridObj);

        // a:tr rows
        for (const row of opts.rows) {
            const tr = new TableRow(row);
            const trObj = tr.prepForXml(context);
            if (trObj) children.push(trObj);
        }

        return { "a:tbl": children };
    }
}
