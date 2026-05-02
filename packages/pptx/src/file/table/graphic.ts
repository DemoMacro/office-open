import { BuilderElement, XmlComponent } from "@file/xml-components";

/**
 * a:graphic > a:graphicData — DrawingML graphic wrapper for table.
 */
export class Graphic extends XmlComponent {
    public constructor(table: XmlComponent) {
        super("a:graphic");
        this.root.push(
            new BuilderElement({
                name: "a:graphicData",
                attributes: {
                    uri: {
                        key: "uri",
                        value: "http://schemas.openxmlformats.org/drawingml/2006/table",
                    },
                },
                children: [table],
            }),
        );
    }
}
