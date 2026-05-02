import { BuilderElement, NextAttributeComponent, XmlComponent } from "@file/xml-components";

export interface IPresentationOptions {
    readonly slideWidth?: number;
    readonly slideHeight?: number;
    readonly slideIds: readonly number[];
}

/**
 * p:presentation — Root element of a PPTX file.
 */
export class Presentation extends XmlComponent {
    public constructor(options: IPresentationOptions) {
        super("p:presentation");
        this.root.push(
            new NextAttributeComponent({
                "xmlns:a": {
                    key: "xmlns:a",
                    value: "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                "xmlns:r": {
                    key: "xmlns:r",
                    value: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                },
                "xmlns:p": {
                    key: "xmlns:p",
                    value: "http://schemas.openxmlformats.org/presentationml/2006/main",
                },
            }),
        );

        this.root.push(
            new BuilderElement({
                name: "p:sldMasterIdLst",
                children: [
                    new BuilderElement({
                        name: "p:sldMasterId",
                        attributes: {
                            id: { key: "id", value: 2147483648 },
                            rId: { key: "r:id", value: "rId1" },
                        },
                    }),
                ],
            }),
        );

        this.root.push(
            new BuilderElement({
                name: "p:sldIdLst",
                children: options.slideIds.map(
                    (_id, i) =>
                        new BuilderElement({
                            name: "p:sldId",
                            attributes: {
                                id: { key: "id", value: 256 + i },
                                rId: { key: "r:id", value: `rId${i + 3}` },
                            },
                        }),
                ),
            }),
        );

        this.root.push(
            new BuilderElement({
                name: "p:sldSz",
                attributes: {
                    cx: { key: "cx", value: options.slideWidth ?? 9144000 },
                    cy: { key: "cy", value: options.slideHeight ?? 6858000 },
                },
            }),
        );

        this.root.push(
            new BuilderElement({
                name: "p:notesSz",
                attributes: {
                    cx: { key: "cx", value: 6858000 },
                    cy: { key: "cy", value: 9144000 },
                },
            }),
        );
    }
}
