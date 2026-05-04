import type { DocPropertiesOptions } from "@file/drawing/doc-properties/doc-properties";
import type { WpsShapeCoreOptions } from "@file/drawing/inline/graphic/graphic-data/wps";

import { Drawing } from "../../drawing";
import type { IFloating } from "../../drawing";
import type { IMediaTransformation, WpsMediaData } from "../../media";
import { createTransformation } from "../../media";
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
            fill: options.fill,
        });

        this.extraChildren.push(drawing);
    }
}
