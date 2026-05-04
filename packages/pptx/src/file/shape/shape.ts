import type { IAnimationOptions } from "@file/animation/types";
import type { IEffectsOptions } from "@file/drawingml/effects";
import type { OutlineOptions } from "@file/drawingml/outline";
import { ShapeProperties } from "@file/drawingml/shape-properties";
import type { IShapePropertiesOptions } from "@file/drawingml/shape-properties";
import type { File } from "@file/file";
import { XmlComponent as Xc } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";
import { pixelsToEmus } from "@util/types";

import { Paragraph } from "./paragraph/paragraph";
import { Run } from "./paragraph/run";
import { TextBody } from "./text-body";
import type { ITextBodyOptions } from "./text-body";

export interface IShapeOptions {
    readonly id?: number;
    readonly name?: string;
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number;
    readonly geometry?: string;
    readonly fill?: IShapePropertiesOptions["fill"];
    readonly outline?: OutlineOptions;
    readonly effects?: IEffectsOptions;
    readonly flipH?: boolean;
    readonly rotation?: number;
    readonly text?: string;
    readonly paragraphs?: ITextBodyOptions["paragraphs"];
    readonly textVertical?: ITextBodyOptions["vertical"];
    readonly textAnchor?: ITextBodyOptions["anchor"];
    readonly textAutoFit?: ITextBodyOptions["autoFit"];
    readonly textWrap?: ITextBodyOptions["wrap"];
    readonly textMargins?: ITextBodyOptions["margins"];
    readonly textColumns?: ITextBodyOptions["columns"];
    readonly textColumnSpacing?: ITextBodyOptions["columnSpacing"];
    readonly animation?: IAnimationOptions;
    readonly placeholder?: "title" | "body" | "subTitle" | "sldNum" | "dt" | "ftr" | "hdr" | "obj";
    readonly placeholderIndex?: number;
}

/**
 * Pure function: builds p:ph element for placeholder.
 */
function buildPlaceholder(type: string, index?: number): IXmlableObject {
    const attrs: Record<string, string | number> = { type };
    if (index !== undefined) attrs.idx = index;
    return { "p:ph": { _attr: attrs } };
}

/**
 * p:sp — A shape on a slide.
 * Lazy: stores options, builds XML object in prepForXml.
 *
 * x/y/width/height accept pixel values and are internally converted to EMUs.
 */
export class Shape extends Xc {
    private static nextId = 2;
    private readonly shapeId: number;
    private readonly animationOptions?: IAnimationOptions;
    private readonly options: IShapeOptions;

    public constructor(options: IShapeOptions = {}) {
        super("p:sp");

        const id = options.id ?? Shape.nextId++;
        this.shapeId = id;
        this.animationOptions = options.animation;
        this.options = { ...options, id };
    }

    public get ShapeId(): number {
        return this.shapeId;
    }

    public get Animation(): IAnimationOptions | undefined {
        return this.animationOptions;
    }

    public override prepForXml(context: IContext): IXmlableObject | undefined {
        const opts = this.options;
        const id = this.shapeId;
        const name = opts.name ?? `Shape ${id}`;
        const children: IXmlableObject[] = [];

        // nvSpPr
        const nvPrChildren: IXmlableObject[] = [];
        if (opts.placeholder) {
            nvPrChildren.push(buildPlaceholder(opts.placeholder, opts.placeholderIndex));
        }
        children.push({
            "p:nvSpPr": [
                { "p:cNvPr": { _attr: { id, name } } },
                { "p:cNvSpPr": {} },
                { "p:nvPr": nvPrChildren.length > 0 ? nvPrChildren : {} },
            ],
        });

        // spPr (ShapeProperties)
        const shapeProps: IShapePropertiesOptions = {
            x: opts.x !== undefined ? pixelsToEmus(opts.x) : undefined,
            y: opts.y !== undefined ? pixelsToEmus(opts.y) : undefined,
            width: opts.width !== undefined ? pixelsToEmus(opts.width) : undefined,
            height: opts.height !== undefined ? pixelsToEmus(opts.height) : undefined,
            geometry: opts.geometry,
            fill: opts.fill,
            outline: opts.outline,
            effects: opts.effects,
            flipH: opts.flipH,
            rotation: opts.rotation,
        };
        const spPr = new ShapeProperties(shapeProps);
        const spPrObj = spPr.prepForXml(context as IContext<File>);
        if (spPrObj) children.push(spPrObj);

        // txBody (TextBody)
        const textBodyOptions: ITextBodyOptions = {
            paragraphs:
                opts.paragraphs ??
                (opts.text
                    ? [new Paragraph({ children: [new Run({ text: opts.text })] })]
                    : undefined),
            vertical: opts.textVertical,
            anchor: opts.textAnchor,
            autoFit: opts.textAutoFit,
            wrap: opts.textWrap,
            margins: opts.textMargins,
            columns: opts.textColumns,
            columnSpacing: opts.textColumnSpacing,
        };
        const txBody = new TextBody(textBodyOptions);
        const txBodyObj = txBody.prepForXml(context);
        if (txBodyObj) children.push(txBodyObj);

        return { "p:sp": children };
    }
}
