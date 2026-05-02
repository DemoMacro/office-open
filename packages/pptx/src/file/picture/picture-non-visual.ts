import { NonVisualPictureProperties } from "@file/drawingml/non-visual-picture-props";
import { BuilderElement, XmlComponent } from "@file/xml-components";

/**
 * p:nvPicPr — Non-visual picture properties for p:pic.
 */
export class PictureNonVisual extends XmlComponent {
    public constructor(id: number, name: string) {
        super("p:nvPicPr");
        this.root.push(
            new BuilderElement({
                name: "a:cNvPr",
                attributes: {
                    id: { key: "id", value: id },
                    name: { key: "name", value: name },
                    descr: { key: "descr", value: "" },
                },
            }),
        );
        this.root.push(new NonVisualPictureProperties());
        this.root.push(new BuilderElement({ name: "p:nvPr" }));
    }
}
