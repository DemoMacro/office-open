import { BuilderElement, XmlComponent } from "@file/xml-components";

import { Paragraph } from "./paragraph/paragraph";
import { Run } from "./paragraph/run";

export interface ITextBodyOptions {
    readonly paragraphs?: readonly (Paragraph | string)[];
    readonly vertical?: "vert" | "vert270" | "horz" | "wordArtVert";
    readonly anchor?: "t" | "ctr" | "b";
    readonly autoFit?: "normal" | "shape" | "none";
    readonly wrap?: "square" | "none";
    readonly margins?: {
        readonly top?: number;
        readonly bottom?: number;
        readonly left?: number;
        readonly right?: number;
    };
    readonly columns?: number;
    readonly columnSpacing?: number;
}

/**
 * p:txBody — Text body within a shape.
 * Uses p: prefix in PresentationML context, though type is a:CT_TextBody.
 */
export class TextBody extends XmlComponent {
    public constructor(options: ITextBodyOptions = {}) {
        super("p:txBody");

        const bodyPrAttrs: Record<string, { readonly key: string; readonly value: string | number }> = {};
        if (options.vertical) bodyPrAttrs.vert = { key: "vert", value: options.vertical };
        if (options.anchor) bodyPrAttrs.anchor = { key: "anchor", value: options.anchor };
        if (options.wrap) bodyPrAttrs.wrap = { key: "wrap", value: options.wrap };
        if (options.margins?.top !== undefined) bodyPrAttrs.tIns = { key: "tIns", value: options.margins.top };
        if (options.margins?.bottom !== undefined) bodyPrAttrs.bIns = { key: "bIns", value: options.margins.bottom };
        if (options.margins?.left !== undefined) bodyPrAttrs.lIns = { key: "lIns", value: options.margins.left };
        if (options.margins?.right !== undefined) bodyPrAttrs.rIns = { key: "rIns", value: options.margins.right };
        if (options.columns !== undefined) bodyPrAttrs.numCol = { key: "numCol", value: options.columns };

        const bodyPrChildren: BuilderElement<{}>[] = [];

        if (options.autoFit === "normal") {
            bodyPrChildren.push(new BuilderElement({ name: "a:normAutofit" }));
        } else if (options.autoFit === "shape") {
            bodyPrChildren.push(new BuilderElement({ name: "a:spAutoFit" }));
        } else if (options.autoFit === "none") {
            bodyPrChildren.push(new BuilderElement({ name: "a:noAutofit" }));
        }

        if (options.columnSpacing !== undefined) {
            bodyPrChildren.push(
                new BuilderElement({
                    name: "a:spcCol",
                    attributes: { spcPts: { key: "spcPts", value: options.columnSpacing * 100 } },
                }),
            );
        }

        this.root.push(
            new BuilderElement({
                name: "a:bodyPr",
                attributes: Object.keys(bodyPrAttrs).length > 0 ? bodyPrAttrs : undefined,
                children: bodyPrChildren.length > 0 ? bodyPrChildren : undefined,
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
