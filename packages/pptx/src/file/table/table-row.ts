import { NextAttributeComponent, XmlComponent } from "@file/xml-components";

import { TableCell, type ITableCellOptions } from "./table-cell";

export interface ITableRowOptions {
    readonly height?: number;
    readonly cells: readonly ITableCellOptions[];
}

/**
 * a:tr — Table row containing cells.
 *
 * Per OOXML spec, `h` is a required attribute on `a:tr`.
 * Children: tc (multiple) → extLst. No trPr in DrawingML tables.
 */
export class TableRow extends XmlComponent {
    public constructor(options: ITableRowOptions) {
        super("a:tr");
        this.root.push(
            new NextAttributeComponent({
                h: { key: "h", value: options.height ?? 0 },
            }),
        );
        for (const cell of options.cells) {
            this.root.push(new TableCell(cell));
        }
    }
}
