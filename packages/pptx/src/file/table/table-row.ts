import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";

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

  public override toXml(context: Context): string {
    const h = this.options.height ?? 0;
    const parts: string[] = [];
    for (const cell of this.options.cells) {
      parts.push(new TableCell(cell).toXml(context));
    }
    return `<a:tr h="${h}">${parts.join("")}</a:tr>`;
  }
}
