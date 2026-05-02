import { BuilderElement, XmlComponent } from "@file/xml-components";

/**
 * a:trPr — Table row properties.
 */
export class TableRowProperties extends XmlComponent {
    public constructor(height?: number) {
        super("a:trPr");
        if (height !== undefined) {
            this.root.push(
                new BuilderElement({
                    name: "a:h",
                    attributes: { val: { key: "val", value: height } },
                }),
            );
        }
    }
}
