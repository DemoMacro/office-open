import { BuilderElement, NextAttributeComponent, XmlComponent } from "@file/xml-components";

export interface IShowOptions {
    readonly loop?: boolean;
    readonly kiosk?: boolean;
    readonly showNarration?: boolean;
    readonly useTimings?: boolean;
}

/**
 * p:presentationPr — Presentation properties.
 */
export class PresentationProperties extends XmlComponent {
    public constructor(showOptions?: IShowOptions) {
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

        if (showOptions) {
            const attrs: Record<string, { readonly key: string; readonly value: number }> = {};
            if (showOptions.loop) attrs.loop = { key: "loop", value: 1 };
            if (showOptions.kiosk) attrs.kiosk = { key: "kiosk", value: 1 };
            if (showOptions.showNarration === false)
                attrs.showNarration = { key: "showNarration", value: 0 };
            if (showOptions.useTimings) attrs.useTimings = { key: "useTimings", value: 1 };

            this.root.push(
                new BuilderElement({
                    name: "p:showPr",
                    children: [new BuilderElement({ name: "p:present" })],
                    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
                }),
            );
        }
    }
}
