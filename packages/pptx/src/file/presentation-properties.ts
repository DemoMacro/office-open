import { NextAttributeComponent, XmlComponent } from "@file/xml-components";

/**
 * p:presentationPr — Presentation properties.
 */
export class PresentationProperties extends XmlComponent {
    public constructor() {
        super("p:presentationPr");
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
    }
}
