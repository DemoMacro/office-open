import { GroupShapeProperties } from "@file/drawingml/group-shape-properties";
import { GroupShapeNonVisualProperties } from "@file/shape-tree/group-shape-non-visual";
import { BuilderElement, NextAttributeComponent, XmlComponent } from "@file/xml-components";

import { Paragraph } from "../shape/paragraph/paragraph";
import { Run } from "../shape/paragraph/run";
import { TextBody } from "../shape/text-body";

export interface INotesSlideOptions {
    readonly text?: string;
}

/**
 * p:notes — A notes slide associated with a presentation slide.
 *
 * Contains a slide image placeholder and a body text area for speaker notes.
 */
export class NotesSlide extends XmlComponent {
    public constructor(options: INotesSlideOptions = {}) {
        super("p:notes");

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

        this.root.push(new NotesCommonSlideData(options.text));
        this.root.push(
            new BuilderElement({
                name: "p:clrMapOvr",
                children: [new BuilderElement({ name: "a:masterClrMapping" })],
            }),
        );
    }
}

class NotesCommonSlideData extends XmlComponent {
    public constructor(text?: string) {
        super("p:cSld");
        this.root.push(new NotesShapeTree(text));
    }
}

class NotesShapeTree extends XmlComponent {
    public constructor(text?: string) {
        super("p:spTree");
        this.root.push(new GroupShapeNonVisualProperties());
        this.root.push(new GroupShapeProperties());
        this.root.push(new SlideImagePlaceholder());
        this.root.push(new NotesBodyPlaceholder(text));
    }
}

class SlideImagePlaceholder extends XmlComponent {
    public constructor() {
        super("p:sp");
        this.root.push(
            new BuilderElement({
                name: "p:nvSpPr",
                children: [
                    new BuilderElement({
                        name: "p:cNvPr",
                        attributes: { id: { key: "id", value: 2 }, name: { key: "name", value: "Slide Image Placeholder 1" } },
                    }),
                    new BuilderElement({ name: "p:cNvSpPr" }),
                    new BuilderElement({
                        name: "p:nvPr",
                        children: [
                            new BuilderElement({
                                name: "p:ph",
                                attributes: { type: { key: "type", value: "sldImg" } },
                            }),
                        ],
                    }),
                ],
            }),
        );
        this.root.push(
            new BuilderElement({
                name: "p:spPr",
                children: [
                    new BuilderElement({
                        name: "a:xfrm",
                        children: [
                            new BuilderElement({
                                name: "a:off",
                                attributes: { x: { key: "x", value: 685800 }, y: { key: "y", value: 1600200 } },
                            }),
                            new BuilderElement({
                                name: "a:ext",
                                attributes: { cx: { key: "cx", value: 5486400 }, cy: { key: "cy", value: 3086100 } },
                            }),
                        ],
                    }),
                ],
            }),
        );
    }
}

class NotesBodyPlaceholder extends XmlComponent {
    public constructor(text?: string) {
        super("p:sp");
        this.root.push(
            new BuilderElement({
                name: "p:nvSpPr",
                children: [
                    new BuilderElement({
                        name: "p:cNvPr",
                        attributes: { id: { key: "id", value: 3 }, name: { key: "name", value: "Notes Placeholder 1" } },
                    }),
                    new BuilderElement({ name: "p:cNvSpPr" }),
                    new BuilderElement({
                        name: "p:nvPr",
                        children: [
                            new BuilderElement({
                                name: "p:ph",
                                attributes: {
                                    type: { key: "type", value: "body" },
                                    idx: { key: "idx", value: 1 },
                                },
                            }),
                        ],
                    }),
                ],
            }),
        );
        this.root.push(
            new BuilderElement({
                name: "p:spPr",
                children: [
                    new BuilderElement({
                        name: "a:xfrm",
                        children: [
                            new BuilderElement({
                                name: "a:off",
                                attributes: { x: { key: "x", value: 685800 }, y: { key: "y", value: 4800600 } },
                            }),
                            new BuilderElement({
                                name: "a:ext",
                                attributes: { cx: { key: "cx", value: 5486400 }, cy: { key: "cy", value: 3600450 } },
                            }),
                        ],
                    }),
                ],
            }),
        );

        const textBodyOptions = text
            ? { paragraphs: [new Paragraph({ children: [new Run({ text })] })] }
            : undefined;
        this.root.push(new TextBody(textBodyOptions));
    }
}
