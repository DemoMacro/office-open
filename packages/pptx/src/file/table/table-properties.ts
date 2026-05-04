import { BaseXmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";

/**
 * a:tblPr — Table properties (firstRow, bandRow, etc.).
 * Lazy: stores options, builds IXmlableObject in prepForXml.
 */
export class TableProperties extends BaseXmlComponent {
    private readonly options?: {
        readonly firstRow?: boolean;
        readonly lastRow?: boolean;
        readonly bandRow?: boolean;
        readonly firstCol?: boolean;
        readonly lastCol?: boolean;
        readonly bandCol?: boolean;
    };

    public constructor(options?: {
        readonly firstRow?: boolean;
        readonly lastRow?: boolean;
        readonly bandRow?: boolean;
        readonly firstCol?: boolean;
        readonly lastCol?: boolean;
        readonly bandCol?: boolean;
    }) {
        super("a:tblPr");
        this.options = options;
    }

    public override prepForXml(_context: IContext): IXmlableObject {
        if (!this.options) return { "a:tblPr": {} };

        const attrs: Record<string, string | number> = {};
        const opts = this.options;
        if (opts.firstRow !== undefined) attrs.firstRow = opts.firstRow ? 1 : 0;
        if (opts.lastRow !== undefined) attrs.lastRow = opts.lastRow ? 1 : 0;
        if (opts.bandRow !== undefined) attrs.bandRow = opts.bandRow ? 1 : 0;
        if (opts.firstCol !== undefined) attrs.firstCol = opts.firstCol ? 1 : 0;
        if (opts.lastCol !== undefined) attrs.lastCol = opts.lastCol ? 1 : 0;
        if (opts.bandCol !== undefined) attrs.bandCol = opts.bandCol ? 1 : 0;

        return { "a:tblPr": Object.keys(attrs).length > 0 ? { _attr: attrs } : {} };
    }
}
