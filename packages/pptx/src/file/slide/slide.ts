import { BuilderElement, NextAttributeComponent, XmlComponent } from "@file/xml-components";

import { CommonSlideData } from "./common-slide-data";

class ColorMapOverride extends BuilderElement<{}> {
    public constructor() {
        super({
            name: "p:clrMapOvr",
            children: [new BuilderElement({ name: "a:masterClrMapping" })],
        });
    }
}

/**
 * p:sld — A slide in a presentation.
 */
export class Slide extends XmlComponent {
    public constructor(children: readonly XmlComponent[]) {
        super("p:sld");
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
        this.root.push(new CommonSlideData(children));
        this.root.push(new ColorMapOverride());
    }
}
