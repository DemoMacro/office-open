/**
 * Table module for WordprocessingML documents.
 *
 * Reference: http://officeopenxml.com/WPtableGrid.php
 *
 * @module
 */
import type { FileChild } from "@file/file-child";
import { BaseXmlComponent } from "@file/xml-components";
import type { Context, IXmlableObject } from "@file/xml-components";

import type { AlignmentType } from "../paragraph";
import { StructuredDocumentTagRow } from "../sdt";
import { TableGrid } from "./grid";
import type { TableGridChangeOptions } from "./grid";
import { TableCell, VerticalMergeType } from "./table-cell";
import type { TableCellSpacingProperties } from "./table-cell-spacing";
import { TableProperties } from "./table-properties";
import type {
  TableBordersOptions,
  TableFloatOptions,
  ITablePropertiesChangeOptions,
} from "./table-properties";
import type { TableCellMarginOptions } from "./table-properties/table-cell-margin";
import type { TableLayoutType } from "./table-properties/table-layout";
import type { TableLookOptions } from "./table-properties/table-look";
import { TableRow } from "./table-row";
import type { TableRowOptions } from "./table-row";
import type { TableWidthProperties } from "./table-width";

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
export interface TableOptions {
  readonly rows: readonly (TableRow | StructuredDocumentTagRow | TableRowOptions)[];
  readonly width?: TableWidthProperties;
  readonly columnWidths?: readonly number[];
  readonly columnWidthsRevision?: TableGridChangeOptions;
  readonly margins?: TableCellMarginOptions;
  readonly indent?: TableWidthProperties;
  readonly float?: TableFloatOptions;
  readonly layout?: (typeof TableLayoutType)[keyof typeof TableLayoutType];
  readonly style?: string;
  readonly borders?: TableBordersOptions;
  readonly alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
  readonly visuallyRightToLeft?: boolean;
  readonly tableLook?: TableLookOptions;
  readonly cellSpacing?: TableCellSpacingProperties;
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

  private readonly options: TableOptions;
  private readonly columnWidths: readonly number[];
  // Coerced rows: plain TableRowOptions are converted to TableRow instances
  private readonly rows: readonly (TableRow | StructuredDocumentTagRow)[];

  public constructor(options: TableOptions) {
    super("w:tbl");
    this.options = options;

    // Coerce plain TableRowOptions objects into TableRow instances
    this.rows = options.rows.map((row) =>
      row instanceof TableRow || row instanceof StructuredDocumentTagRow ? row : new TableRow(row),
    );

    this.columnWidths =
      options.columnWidths ??
      Array<number>(Math.max(...this.rows.map((row) => row.cellCount))).fill(100);

    // Register CONTINUE cells on subsequent rows for vertical merge
    this.rows.forEach((row, rowIndex) => {
      if (rowIndex === this.rows.length - 1) return;
      if (!(row instanceof TableRow)) return;

      let columnIndex = 0;
      row.cells.forEach((cell) => {
        if (cell.options.rowSpan && cell.options.rowSpan > 1) {
          const nextRow = this.rows[rowIndex + 1];
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

  public prepForXml(context: Context): IXmlableObject | undefined {
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

    const gridObj = new TableGrid(this.columnWidths, this.options.columnWidthsRevision).prepForXml(
      context,
    );
    if (gridObj) children.push(gridObj);

    for (const row of this.rows) {
      const obj = row.prepForXml(context);
      if (obj) children.push(obj);
    }

    return { "w:tbl": children };
  }
}
