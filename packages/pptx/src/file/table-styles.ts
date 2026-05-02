import { NextAttributeComponent, XmlComponent } from "@file/xml-components";

/**
 * a:tblStyleLst — Table style list (required when presentation contains tables).
 */
export class TableStyles extends XmlComponent {
    public constructor() {
        super("a:tblStyleLst");
        this.root.push(
            new NextAttributeComponent({
                "xmlns:a": {
                    key: "xmlns:a",
                    value: "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                def: { key: "def", value: "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" },
            }),
        );
    }
}
