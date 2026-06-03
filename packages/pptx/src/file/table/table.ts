import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";

import type { CellBorderOptions } from "./table-cell-properties";
import { TableGrid } from "./table-grid";
import { TableProperties } from "./table-properties";
import { TableRow, type TableRowOptions } from "./table-row";

export interface DrawingTableOptions {
  readonly rows: readonly TableRowOptions[];
  readonly columnWidths?: readonly number[];
  readonly firstRow?: boolean;
  readonly lastRow?: boolean;
  readonly bandRow?: boolean;
  readonly firstCol?: boolean;
  readonly lastCol?: boolean;
  readonly bandCol?: boolean;
  readonly tableStyleId?: string;
  readonly borders?: {
    readonly top?: CellBorderOptions;
    readonly bottom?: CellBorderOptions;
    readonly left?: CellBorderOptions;
    readonly right?: CellBorderOptions;
  };
}

/**
 * a:tbl — DrawingML table element.
 * Lazy: stores options, builds XML in toXml().
 */
export class DrawingTable extends BaseXmlComponent {
  private readonly options: DrawingTableOptions;

  public constructor(options: DrawingTableOptions) {
    super("a:tbl");
    this.options = options;
  }

  public override toXml(context: Context): string {
    const opts = this.options;
    const parts: string[] = [];

    // a:tblPr
    const tblPr = new TableProperties({
      firstRow: opts.firstRow,
      lastRow: opts.lastRow,
      bandRow: opts.bandRow,
      firstCol: opts.firstCol,
      lastCol: opts.lastCol,
      bandCol: opts.bandCol,
      tableStyleId: opts.tableStyleId,
    });
    parts.push(tblPr.toXml(context));

    // a:tblGrid
    const colWidths =
      opts.columnWidths && opts.columnWidths.length > 0
        ? opts.columnWidths
        : Array.from({ length: opts.rows[0]?.cells.length ?? 1 }, () => 0);
    parts.push(new TableGrid(colWidths).toXml(context));

    // a:tr — distribute table-level borders to edge cells
    const rowCount = opts.rows.length;
    for (let ri = 0; ri < rowCount; ri++) {
      const row = opts.rows[ri];
      const cells = this.distributeBorders(row, ri, rowCount, opts.borders);
      parts.push(new TableRow({ ...row, cells }).toXml(context));
    }

    return `<a:tbl>${parts.join("")}</a:tbl>`;
  }

  /** Distribute table-level borders to edge cells only when needed. */
  private distributeBorders(
    row: TableRowOptions,
    ri: number,
    rowCount: number,
    tb: DrawingTableOptions["borders"],
  ): TableRowOptions["cells"] {
    if (!tb) return row.cells;
    const colCount = row.cells.length;
    return row.cells.map((cell, ci) => {
      const needTop = ri === 0 && !!tb.top && !cell.borders?.top;
      const needBottom = ri === rowCount - 1 && !!tb.bottom && !cell.borders?.bottom;
      const needLeft = ci === 0 && !!tb.left && !cell.borders?.left;
      const needRight = ci === colCount - 1 && !!tb.right && !cell.borders?.right;
      if (!needTop && !needBottom && !needLeft && !needRight) return cell;
      const borders = {
        ...cell.borders,
        ...(needTop && { top: tb.top }),
        ...(needBottom && { bottom: tb.bottom }),
        ...(needLeft && { left: tb.left }),
        ...(needRight && { right: tb.right }),
      };
      return { ...cell, borders };
    });
  }
}
