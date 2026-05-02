import { BuilderElement } from "@file/xml-components";

/**
 * a:cNvPr — Non-visual drawing properties (id, name).
 */
export class NonVisualDrawingProperties extends BuilderElement<{
    readonly id: number;
    readonly name: string;
}> {
    public constructor(id: number, name: string) {
        super({
            name: "a:cNvPr",
            attributes: { id: { key: "id", value: id }, name: { key: "name", value: name } },
        });
    }
}
