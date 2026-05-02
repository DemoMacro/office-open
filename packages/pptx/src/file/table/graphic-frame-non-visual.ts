import { BuilderElement, XmlComponent } from "@file/xml-components";

/**
 * p:nvGraphicFramePr — Non-visual properties for graphic frame.
 */
export class GraphicFrameNonVisual extends XmlComponent {
    private static nextId = 1024;

    public constructor() {
        super("p:nvGraphicFramePr");
        const id = GraphicFrameNonVisual.nextId++;
        this.root.push(
            new BuilderElement({
                name: "p:cNvPr",
                attributes: {
                    id: { key: "id", value: id },
                    name: { key: "name", value: `Table ${id}` },
                },
            }),
        );
        this.root.push(
            new BuilderElement({
                name: "p:cNvGraphicFramePr",
                children: [
                    new BuilderElement({
                        name: "a:graphicFrameLocks",
                        attributes: { noGrp: { key: "noGrp", value: 1 } },
                    }),
                ],
            }),
        );
        this.root.push(new BuilderElement({ name: "p:nvPr" }));
    }
}
