import { BuilderElement } from "@file/xml-components";

/**
 * p:nvGrpSpPr — Non-visual properties for the shape tree (group shape).
 */
export class GroupShapeNonVisualProperties extends BuilderElement<{}> {
    public constructor() {
        super({
            name: "p:nvGrpSpPr",
            children: [
                new BuilderElement({
                    name: "p:cNvPr",
                    attributes: { id: { key: "id", value: 1 }, name: { key: "name", value: "" } },
                }),
                new BuilderElement({ name: "p:cNvGrpSpPr" }),
                new BuilderElement({ name: "p:nvPr" }),
            ],
        });
    }
}
