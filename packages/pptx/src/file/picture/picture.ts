import { BlipFill } from "@file/drawingml/blip-fill";
import { PresetGeometry } from "@file/drawingml/preset-geometry";
import { Transform2D } from "@file/drawingml/transform-2d";
import type { File } from "@file/file";
import type { IMediaData } from "@file/media/data";
import { BuilderElement, type IContext, XmlComponent } from "@file/xml-components";
import { pixelsToEmus } from "@util/types";

import { PictureNonVisual } from "./picture-non-visual";

export interface IPictureOptions {
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number;
    readonly data: Uint8Array;
    readonly type: "png" | "jpg" | "gif" | "bmp" | "emf" | "wmf";
    readonly name?: string;
}

/**
 * p:pic — A picture on a slide.
 *
 * Registers image with Media collection via prepForXml.
 * The ImageReplacer replaces `{fileName}` placeholder with actual rId.
 */
export class Picture extends XmlComponent {
    private static nextId = 100;
    private readonly imageData: IMediaData;

    public constructor(options: IPictureOptions) {
        super("p:pic");

        const id = Picture.nextId++;
        const name = options.name ?? `Picture ${id}`;
        const fileName = `${name.replace(/\s+/g, "_")}.${options.type}`;

        this.imageData = {
            type: options.type,
            fileName,
            transformation: {
                pixels: {
                    x: options.width ?? 0,
                    y: options.height ?? 0,
                },
                emus: {
                    x: pixelsToEmus(options.width ?? 0),
                    y: pixelsToEmus(options.height ?? 0),
                },
            },
            data: options.data,
        };

        this.root.push(new PictureNonVisual(id, name));

        this.root.push(new BlipFill(fileName));

        this.root.push(
            new BuilderElement({
                name: "p:spPr",
                children: [
                    new Transform2D({
                        x: pixelsToEmus(options.x ?? 0),
                        y: pixelsToEmus(options.y ?? 0),
                        width: pixelsToEmus(options.width ?? 0),
                        height: pixelsToEmus(options.height ?? 0),
                    }),
                    new PresetGeometry("rect"),
                ],
            }),
        );
    }

    public override prepForXml(context: IContext) {
        (context.fileData as File)?.Media.addImage(this.imageData.fileName, this.imageData);
        return super.prepForXml(context);
    }
}
