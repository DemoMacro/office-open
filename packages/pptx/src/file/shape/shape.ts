import { NonVisualShapeProperties } from "@file/drawingml/non-visual-shape-props";
import { Outline, type OutlineOptions } from "@file/drawingml/outline";
import { ShapeProperties } from "@file/drawingml/shape-properties";
import type { IShapePropertiesOptions } from "@file/drawingml/shape-properties";
import type { IEffectsOptions } from "@file/drawingml/effects";
import { BuilderElement, XmlComponent as Xc } from "@file/xml-components";
import { pixelsToEmus } from "@util/types";
import type { IAnimationOptions } from "@file/animation/types";

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
}

/**
 * p:sp — A shape on a slide.
 *
 * x/y/width/height accept pixel values and are internally converted to EMUs.
 */
export class Shape extends Xc {
    private static nextId = 2;
    private readonly shapeId: number;
    private readonly animationOptions?: IAnimationOptions;

    public constructor(options: IShapeOptions = {}) {
        super("p:sp");

        const id = options.id ?? Shape.nextId++;
        this.shapeId = id;
        this.animationOptions = options.animation;
        const name = options.name ?? `Shape ${id}`;

        this.root.push(
            new BuilderElement({
                name: "p:nvSpPr",
                children: [
                    new BuilderElement({
                        name: "p:cNvPr",
                        attributes: {
                            id: { key: "id", value: id },
                            name: { key: "name", value: name },
                        },
                    }),
                    new NonVisualShapeProperties(),
                    new BuilderElement({
                        name: "p:nvPr",
                    }),
                ],
            }),
        );

        const shapeProps: IShapePropertiesOptions = {
            x: options.x !== undefined ? pixelsToEmus(options.x) : undefined,
            y: options.y !== undefined ? pixelsToEmus(options.y) : undefined,
            width: options.width !== undefined ? pixelsToEmus(options.width) : undefined,
            height: options.height !== undefined ? pixelsToEmus(options.height) : undefined,
            geometry: options.geometry,
            fill: options.fill,
            outline: options.outline,
            effects: options.effects,
            flipH: options.flipH,
            rotation: options.rotation,
        };
        this.root.push(new ShapeProperties(shapeProps));

        const textBodyOptions: ITextBodyOptions = {
            paragraphs:
                options.paragraphs ??
                (options.text
                    ? [new Paragraph({ children: [new Run({ text: options.text })] })]
                    : undefined),
            vertical: options.textVertical,
            anchor: options.textAnchor,
            autoFit: options.textAutoFit,
            wrap: options.textWrap,
            margins: options.textMargins,
            columns: options.textColumns,
            columnSpacing: options.textColumnSpacing,
        };
        this.root.push(new TextBody(textBodyOptions));
    }

    public get ShapeId(): number {
        return this.shapeId;
    }

    public get Animation(): IAnimationOptions | undefined {
        return this.animationOptions;
    }
}
