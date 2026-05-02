import { BuilderElement } from "@file/xml-components";

/**
 * p:nvGrpSpPr — Non-visual properties for a group shape.
 */
export class GroupShapeNonVisualProperties extends BuilderElement<{}> {
    public constructor(id: number = 1, name: string = "") {
        super({
            name: "p:nvGrpSpPr",
            children: [
                new BuilderElement({
                    name: "p:cNvPr",
                    attributes: { id: { key: "id", value: id }, name: { key: "name", value: name } },
                }),
                new BuilderElement({ name: "p:cNvGrpSpPr" }),
                new BuilderElement({ name: "p:nvPr" }),
            ],
        });
    }
}
