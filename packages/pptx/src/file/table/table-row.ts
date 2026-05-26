import { BaseXmlComponent } from "@file/xml-components";
import type { Context, IXmlableObject } from "@file/xml-components";

import { TableCell, type TableCellOptions } from "./table-cell";

export interface TableRowOptions {
  readonly height?: number;
  readonly cells: readonly TableCellOptions[];
}

/**
 * a:tr — Table row containing cells.
 * Lazy: stores options, builds IXmlableObject in prepForXml.
 */
export class TableRow extends BaseXmlComponent {
  private readonly options: TableRowOptions;

  public constructor(options: TableRowOptions) {
    super("a:tr");
    this.options = options;
  }

  public override prepForXml(context: Context): IXmlableObject {
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
