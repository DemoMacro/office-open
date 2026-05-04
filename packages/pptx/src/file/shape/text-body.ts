import { XmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";

import { VerticalAlignment } from "../table/table-cell";
import { Paragraph } from "./paragraph/paragraph";
import { Run } from "./paragraph/run";

export interface ITextBodyOptions {
    readonly paragraphs?: readonly (Paragraph | string)[];
    readonly vertical?: "vert" | "vert270" | "horz" | "wordArtVert";
    readonly anchor?: keyof typeof VerticalAlignment;
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
 * Pure function: builds a:bodyPr content.
 */
function buildBodyPr(options: ITextBodyOptions): IXmlableObject {
    const bodyPrChildren: IXmlableObject[] = [];

    if (options.autoFit === "normal") bodyPrChildren.push({ "a:normAutofit": {} });
    else if (options.autoFit === "shape") bodyPrChildren.push({ "a:spAutoFit": {} });
    else if (options.autoFit === "none") bodyPrChildren.push({ "a:noAutofit": {} });

    const bodyPrContent: IXmlableObject[] = [];

    const attrs: Record<string, string | number> = {};
    if (options.vertical) attrs.vert = options.vertical;
    if (options.anchor) attrs.anchor = VerticalAlignment[options.anchor];
    if (options.wrap) attrs.wrap = options.wrap;
    if (options.margins?.top !== undefined) attrs.tIns = options.margins.top;
    if (options.margins?.bottom !== undefined) attrs.bIns = options.margins.bottom;
    if (options.margins?.left !== undefined) attrs.lIns = options.margins.left;
    if (options.margins?.right !== undefined) attrs.rIns = options.margins.right;
    if (options.columns !== undefined) attrs.numCol = options.columns;
    if (options.columnSpacing !== undefined) attrs.spcCol = options.columnSpacing * 100;
    if (Object.keys(attrs).length > 0) bodyPrContent.push({ _attr: attrs });

    bodyPrContent.push(...bodyPrChildren);

    return {
        "a:bodyPr":
            bodyPrContent.length === 1 && "_attr" in bodyPrContent[0]
                ? bodyPrContent[0]
                : bodyPrContent.length > 0
                  ? bodyPrContent
                  : {},
    };
}

/**
 * p:txBody — Text body within a shape.
 * Lazy: stores options, builds XML object in prepForXml.
 */
export class TextBody extends XmlComponent {
    private readonly options: ITextBodyOptions;

    public constructor(options: ITextBodyOptions = {}) {
        super("p:txBody");
        this.options = options;
    }

    public prepForXml(context: IContext): IXmlableObject | undefined {
        const children: IXmlableObject[] = [];

        children.push(buildBodyPr(this.options));
        children.push({ "a:lstStyle": {} });

        if (this.options.paragraphs) {
            for (const p of this.options.paragraphs) {
                const para =
                    typeof p === "string" ? new Paragraph({ children: [new Run({ text: p })] }) : p;
                const obj = para.prepForXml(context);
                if (obj) children.push(obj);
            }
        } else {
            const obj = new Paragraph().prepForXml(context);
            if (obj) children.push(obj);
        }

        return { "p:txBody": children };
    }
}
