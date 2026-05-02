import { BuilderElement, NextAttributeComponent, XmlComponent } from "@file/xml-components";

/**
 * p:viewPr — View properties.
 */
export class ViewProperties extends XmlComponent {
    public constructor() {
        super("p:viewPr");
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
                name: "p:normalViewPr",
                children: [
                    new BuilderElement({
                        name: "p:restoredLeft",
                        attributes: {
                            sz: { key: "sz", value: 14996 },
                            autoAdjust: { key: "autoAdjust", value: 0 },
                        },
                    }),
                    new BuilderElement({
                        name: "p:restoredTop",
                        attributes: { sz: { key: "sz", value: 94660 } },
                    }),
                ],
            }),
        );

        this.root.push(
            new BuilderElement({
                name: "p:slideViewPr",
                children: [
                    new BuilderElement({
                        name: "p:cSldViewPr",
                        attributes: { snapToGrid: { key: "snapToGrid", value: 0 } },
                        children: [
                            new BuilderElement({
                                name: "p:cViewPr",
                                attributes: { varScale: { key: "varScale", value: 1 } },
                                children: [
                                    new BuilderElement({
                                        name: "p:scale",
                                        children: [
                                            new BuilderElement({
                                                name: "a:sx",
                                                attributes: {
                                                    n: { key: "n", value: 90 },
                                                    d: { key: "d", value: 100 },
                                                },
                                            }),
                                            new BuilderElement({
                                                name: "a:sy",
                                                attributes: {
                                                    n: { key: "n", value: 90 },
                                                    d: { key: "d", value: 100 },
                                                },
                                            }),
                                        ],
                                    }),
                                    new BuilderElement({
                                        name: "p:origin",
                                        attributes: {
                                            x: { key: "x", value: 1200 },
                                            y: { key: "y", value: 72 },
                                        },
                                    }),
                                ],
                            }),
                            new BuilderElement({ name: "p:guideLst" }),
                        ],
                    }),
                ],
            }),
        );

        this.root.push(
            new BuilderElement({
                name: "p:notesTextViewPr",
                children: [
                    new BuilderElement({
                        name: "p:cViewPr",
                        children: [
                            new BuilderElement({
                                name: "p:scale",
                                children: [
                                    new BuilderElement({
                                        name: "a:sx",
                                        attributes: {
                                            n: { key: "n", value: 1 },
                                            d: { key: "d", value: 1 },
                                        },
                                    }),
                                    new BuilderElement({
                                        name: "a:sy",
                                        attributes: {
                                            n: { key: "n", value: 1 },
                                            d: { key: "d", value: 1 },
                                        },
                                    }),
                                ],
                            }),
                            new BuilderElement({
                                name: "p:origin",
                                attributes: {
                                    x: { key: "x", value: 0 },
                                    y: { key: "y", value: 0 },
                                },
                            }),
                        ],
                    }),
                ],
            }),
        );

        this.root.push(
            new BuilderElement({
                name: "p:gridSpacing",
                attributes: { cx: { key: "cx", value: 72008 }, cy: { key: "cy", value: 72008 } },
            }),
        );
    }
}
