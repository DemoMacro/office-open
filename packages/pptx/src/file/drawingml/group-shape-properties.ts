import { BuilderElement, XmlComponent } from "@file/xml-components";

/**
 * p:grpSpPr — Group shape properties (transform for shape tree).
 * Uses p: prefix in PresentationML context, though type is a:CT_GroupShapeProperties.
 */
export class GroupShapeProperties extends XmlComponent {
    public constructor() {
        super("p:grpSpPr");
        this.root.push(
            new BuilderElement({
                name: "a:xfrm",
                children: [
                    new BuilderElement({
                        name: "a:off",
                        attributes: {
                            x: { key: "x", value: 0 },
                            y: { key: "y", value: 0 },
                        },
                    }),
                    new BuilderElement({
                        name: "a:ext",
                        attributes: {
                            cx: { key: "cx", value: 0 },
                            cy: { key: "cy", value: 0 },
                        },
                    }),
                    new BuilderElement({
                        name: "a:chOff",
                        attributes: {
                            x: { key: "x", value: 0 },
                            y: { key: "y", value: 0 },
                        },
                    }),
                    new BuilderElement({
                        name: "a:chExt",
                        attributes: {
                            cx: { key: "cx", value: 0 },
                            cy: { key: "cy", value: 0 },
                        },
                    }),
                ],
            }),
        );
    }
}
