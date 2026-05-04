import { BaseXmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";

import { TableCell, type ITableCellOptions } from "./table-cell";

export interface ITableRowOptions {
    readonly height?: number;
    readonly cells: readonly ITableCellOptions[];
}

/**
 * a:tr — Table row containing cells.
 * Lazy: stores options, builds IXmlableObject in prepForXml.
 */
export class TableRow extends BaseXmlComponent {
    private readonly options: ITableRowOptions;

    public constructor(options: ITableRowOptions) {
        super("a:tr");
        this.options = options;
    }

    public override prepForXml(context: IContext): IXmlableObject {
        const children: IXmlableObject[] = [];
        children.push({ _attr: { h: this.options.height ?? 0 } });

        for (const cell of this.options.cells) {
            const tc = new TableCell(cell);
            const obj = tc.prepForXml(context);
            if (obj) children.push(obj);
        }

        return { "a:tr": children };
    }
}
