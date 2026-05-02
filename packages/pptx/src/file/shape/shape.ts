import { NonVisualShapeProperties } from "@file/drawingml/non-visual-shape-props";
import { Outline } from "@file/drawingml/outline";
import { ShapeProperties } from "@file/drawingml/shape-properties";
import type { IShapePropertiesOptions } from "@file/drawingml/shape-properties";
import { BuilderElement, XmlComponent as Xc } from "@file/xml-components";
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
    readonly outline?: Outline;
    readonly flipH?: boolean;
    readonly rotation?: number;
    readonly text?: string;
    readonly paragraphs?: ITextBodyOptions["paragraphs"];
}

/**
 * p:sp — A shape on a slide.
 *
 * x/y/width/height accept pixel values and are internally converted to EMUs.
 */
export class Shape extends Xc {
    private static nextId = 2;

    public constructor(options: IShapeOptions = {}) {
        super("p:sp");

        const id = options.id ?? Shape.nextId++;
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
        };
        this.root.push(new TextBody(textBodyOptions));
    }
}
