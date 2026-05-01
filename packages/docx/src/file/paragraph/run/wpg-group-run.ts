import type { DocPropertiesOptions } from "@file/drawing/doc-properties/doc-properties";
import type { IContext, IXmlableObject } from "@file/xml-components";

import { Run } from ".";
import { Drawing } from "../../drawing";
import type { IFloating } from "../../drawing";
import type {
    IGroupChildMediaData,
    IMediaData,
    IMediaTransformation,
    WpgMediaData,
} from "../../media";
import { createTransformation } from "../../media";

export * from "@file/drawing/inline/graphic/graphic-data/wps/body-properties";

interface CoreGroupOptions {
    readonly children: readonly IGroupChildMediaData[];
    readonly transformation: IMediaTransformation;
    readonly floating?: IFloating;
    readonly altText?: DocPropertiesOptions;
}

/**
 * @publicApi
 */
export type IWpgGroupOptions = { readonly type: "wpg" } & CoreGroupOptions;

/**
 * @publicApi
 */
export class WpgGroupRun extends Run {
    private readonly wpgGroupData: WpgMediaData;
    private readonly mediaDatas: readonly IMediaData[];

    public constructor(options: IWpgGroupOptions) {
        super({});

        this.wpgGroupData = {
            children: options.children,
            transformation: createTransformation(options.transformation),
            type: options.type,
        };
        const drawing = new Drawing(this.wpgGroupData, {
            docProperties: options.altText,
            floating: options.floating,
        });

        this.mediaDatas = options.children
            .filter((child) => child.type !== "wps")
            .map((child) => child as IMediaData);

        this.root.push(drawing);
    }

    public prepForXml(context: IContext): IXmlableObject | undefined {
        this.mediaDatas.forEach((child) => {
            context.file.Media.addImage(child.fileName, child);

            if (child.type === "svg") {
                context.file.Media.addImage(child.fallback.fileName, child.fallback);
            }
        });
        return super.prepForXml(context);
    }
}
