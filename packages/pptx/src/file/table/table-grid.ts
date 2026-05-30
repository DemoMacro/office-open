import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";

/**
 * a:tblGrid — Table grid with column width definitions.
 * Lazy: stores widths, builds IXmlableObject in prepForXml.
 */
export class TableGrid extends BaseXmlComponent {
  private readonly columnWidths: readonly number[];

  public constructor(columnWidths: readonly number[]) {
    super("a:tblGrid");
    this.columnWidths = columnWidths;
  }

  public override toXml(_context: Context): string {
    const cols = this.columnWidths.map((w) => `<a:gridCol w="${w}"/>`).join("");
    return `<a:tblGrid>${cols}</a:tblGrid>`;
  }
}
