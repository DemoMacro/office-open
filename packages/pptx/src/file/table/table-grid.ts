import { BuilderElement, XmlComponent } from "@file/xml-components";

/**
 * a:tblGrid — Table grid with column width definitions.
 */
export class TableGrid extends XmlComponent {
    public constructor(columnWidths: readonly number[]) {
        super("a:tblGrid");
        for (const width of columnWidths) {
            this.root.push(
                new BuilderElement({
                    name: "a:gridCol",
                    attributes: { w: { key: "w", value: width } },
                }),
            );
        }
    }
}
