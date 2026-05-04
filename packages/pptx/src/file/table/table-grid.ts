import { BaseXmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";

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

    public override prepForXml(_context: IContext): IXmlableObject {
        const children = this.columnWidths.map((w) => ({
            "a:gridCol": { _attr: { w } },
        }));
        return { "a:tblGrid": children };
    }
}
