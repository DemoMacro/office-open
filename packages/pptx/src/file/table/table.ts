import { BaseXmlComponent } from "@file/xml-components";
import type { Context, IXmlableObject } from "@file/xml-components";

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
  readonly borders?: {
    readonly top?: CellBorderOptions;
    readonly bottom?: CellBorderOptions;
    readonly left?: CellBorderOptions;
    readonly right?: CellBorderOptions;
  };
}

/**
 * a:tbl — DrawingML table element.
 * Lazy: stores options, builds IXmlableObject in prepForXml.
 */
export class DrawingTable extends BaseXmlComponent {
  private readonly options: DrawingTableOptions;

  public constructor(options: DrawingTableOptions) {
    super("a:tbl");
    this.options = options;
  }

  public override prepForXml(context: Context): IXmlableObject {
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

    // a:tr rows — distribute table-level borders to edge cells
    const tb = opts.borders;
    const rowCount = opts.rows.length;
    for (let ri = 0; ri < rowCount; ri++) {
      const row = opts.rows[ri];
      const colCount = row.cells.length;
      const cells = tb
        ? row.cells.map((cell, ci) => {
            const b = { ...cell.borders };
            if (ri === 0 && tb.top && !b.top) b.top = tb.top;
            if (ri === rowCount - 1 && tb.bottom && !b.bottom) b.bottom = tb.bottom;
            if (ci === 0 && tb.left && !b.left) b.left = tb.left;
            if (ci === colCount - 1 && tb.right && !b.right) b.right = tb.right;
            return Object.keys(b).length === 0 ? cell : { ...cell, borders: b };
          })
        : row.cells;
      const tr = new TableRow({ ...row, cells });
      const trObj = tr.prepForXml(context);
      if (trObj) children.push(trObj);
    }

    return { "a:tbl": children };
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
    });
    parts.push(tblPr.toXml(context));

    // a:tblGrid
    const colWidths =
      opts.columnWidths && opts.columnWidths.length > 0
        ? [...opts.columnWidths]
        : Array.from({ length: opts.rows[0]?.cells.length ?? 1 }, () => 0);
    parts.push(new TableGrid(colWidths).toXml(context));

    // a:tr — distribute table-level borders to edge cells
    const tb = opts.borders;
    const rowCount = opts.rows.length;
    for (let ri = 0; ri < rowCount; ri++) {
      const row = opts.rows[ri];
      const colCount = row.cells.length;
      const cells = tb
        ? row.cells.map((cell, ci) => {
            const b = { ...cell.borders };
            if (ri === 0 && tb.top && !b.top) b.top = tb.top;
            if (ri === rowCount - 1 && tb.bottom && !b.bottom) b.bottom = tb.bottom;
            if (ci === 0 && tb.left && !b.left) b.left = tb.left;
            if (ci === colCount - 1 && tb.right && !b.right) b.right = tb.right;
            return Object.keys(b).length === 0 ? cell : { ...cell, borders: b };
          })
        : row.cells;
      parts.push(new TableRow({ ...row, cells }).toXml(context));
    }

    return `<a:tbl>${parts.join("")}</a:tbl>`;
  }
}
