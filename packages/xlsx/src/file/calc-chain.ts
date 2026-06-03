/**
 * Calculation Chain — generates xl/calcChain.xml.
 *
 * The calculation chain lists formula cells in calculation order,
 * enabling faster recalculation in spreadsheet applications.
 *
 * Reference: OOXML transitional, sml.xsd, CT_CalcChain / CT_CalcCell
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { attrs } from "@office-open/xml";

export interface CalcCell {
  /** Cell reference, e.g. "A1" */
  readonly r: string;
  /** Sheet index (1-based) */
  readonly i: number;
  /** Array formula */
  readonly a?: boolean;
}

export class CalcChain extends BaseXmlComponent {
  private readonly cells: CalcCell[] = [];

  public constructor() {
    super("calcChain");
  }

  public addCell(cell: CalcCell): void {
    this.cells.push(cell);
  }

  public get count(): number {
    return this.cells.length;
  }

  public override toXml(_context: Context): string {
    const parts: string[] = [
      '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    ];
    for (const cell of this.cells) {
      const cellAttrs: Record<string, string | number | boolean> = { r: cell.r, i: cell.i };
      if (cell.a) cellAttrs.a = true;
      parts.push(`<c${attrs(cellAttrs)}/>`);
    }
    parts.push("</calcChain>");
    return parts.join("");
  }
}
