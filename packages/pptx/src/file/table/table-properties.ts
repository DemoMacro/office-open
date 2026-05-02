import { NextAttributeComponent, XmlComponent } from "@file/xml-components";

/**
 * a:tblPr — Table properties (firstRow, bandRow, etc.).
 */
export class TableProperties extends XmlComponent {
    public constructor(options?: {
        readonly firstRow?: boolean;
        readonly lastRow?: boolean;
        readonly bandRow?: boolean;
        readonly firstCol?: boolean;
        readonly lastCol?: boolean;
        readonly bandCol?: boolean;
    }) {
        super("a:tblPr");

        if (!options) return;

        const attrs: Record<
            string,
            { readonly key: string; readonly value: string | number | boolean | undefined }
        > = {};
        if (options.firstRow !== undefined)
            attrs.firstRow = { key: "firstRow", value: options.firstRow ? 1 : 0 };
        if (options.lastRow !== undefined)
            attrs.lastRow = { key: "lastRow", value: options.lastRow ? 1 : 0 };
        if (options.bandRow !== undefined)
            attrs.bandRow = { key: "bandRow", value: options.bandRow ? 1 : 0 };
        if (options.firstCol !== undefined)
            attrs.firstCol = { key: "firstCol", value: options.firstCol ? 1 : 0 };
        if (options.lastCol !== undefined)
            attrs.lastCol = { key: "lastCol", value: options.lastCol ? 1 : 0 };
        if (options.bandCol !== undefined)
            attrs.bandCol = { key: "bandCol", value: options.bandCol ? 1 : 0 };

        if (Object.keys(attrs).length > 0) {
            this.root.push(new NextAttributeComponent(attrs));
        }
    }
}
