import { BuilderElement, XmlComponent } from "@file/xml-components";

import { Paragraph } from "./paragraph/paragraph";
import { Run } from "./paragraph/run";

export interface ITextBodyOptions {
    readonly paragraphs?: readonly (Paragraph | string)[];
}

/**
 * p:txBody — Text body within a shape.
 * Uses p: prefix in PresentationML context, though type is a:CT_TextBody.
 */
export class TextBody extends XmlComponent {
    public constructor(options: ITextBodyOptions = {}) {
        super("p:txBody");

        this.root.push(
            new BuilderElement({
                name: "a:bodyPr",
            }),
        );

        this.root.push(
            new BuilderElement({
                name: "a:lstStyle",
            }),
        );

        if (options.paragraphs) {
            for (const p of options.paragraphs) {
                this.root.push(
                    typeof p === "string" ? new Paragraph({ children: [new Run({ text: p })] }) : p,
                );
            }
        } else {
            this.root.push(new Paragraph());
        }
    }
}
