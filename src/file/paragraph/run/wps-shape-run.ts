import type { DocPropertiesOptions } from "@file/drawing/doc-properties/doc-properties";
import type { WpsShapeCoreOptions } from "@file/drawing/inline/graphic/graphic-data/wps";

import { Drawing } from "../../drawing";
import type { IFloating } from "../../drawing";
import type { IMediaDataTransformation, IMediaTransformation, WpsMediaData } from "../../media";
import { Run } from "../run";

export * from "@file/drawing/inline/graphic/graphic-data/wps/body-properties";

interface CoreShapeOptions {
    readonly transformation: IMediaTransformation;
    readonly floating?: IFloating;
    readonly altText?: DocPropertiesOptions;
}

/**
 * @publicApi
 */
export type IWpsShapeOptions = WpsShapeCoreOptions & { readonly type: "wps" } & CoreShapeOptions;

export const createTransformation = (options: IMediaTransformation): IMediaDataTransformation => ({
    emus: {
        x: Math.round(options.width * 9525),
        y: Math.round(options.height * 9525),
    },
    flip: options.flip,
    offset: {
        emus: {
            x: Math.round((options.offset?.left ?? 0) * 9525),
            y: Math.round((options.offset?.top ?? 0) * 9525),
        },
        pixels: {
            x: Math.round(options.offset?.left ?? 0),
            y: Math.round(options.offset?.top ?? 0),
        },
    },
    pixels: {
        x: Math.round(options.width),
        y: Math.round(options.height),
    },
    rotation: options.rotation ? options.rotation * 60_000 : undefined,
});

/**
 * @publicApi
 */
export class WpsShapeRun extends Run {
    private readonly wpsShapeData: WpsMediaData;

    public constructor(options: IWpsShapeOptions) {
        super({});

        this.wpsShapeData = {
            data: { ...options },
            transformation: createTransformation(options.transformation),
            type: options.type,
        };
        const drawing = new Drawing(this.wpsShapeData, {
            docProperties: options.altText,
            floating: options.floating,
            outline: options.outline,
            solidFill: options.solidFill,
        });

        this.root.push(drawing);
    }
}
