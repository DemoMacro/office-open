import type { File } from "@file/file";
import {
    BuilderElement,
    type IContext,
    type IXmlableObject,
    XmlComponent as Xc,
} from "@file/xml-components";

import { EffectList, buildScene3D, buildShape3D, type IEffectsOptions } from "./effects";
import { buildFill, extractBlipFillMedia } from "./fill";
import type { FillOptions } from "./fill";
import { createOutlineCompat } from "./outline";
import type { OutlineOptions } from "./outline";
import { PresetGeometry } from "./preset-geometry";
import type { ITransform2DOptions } from "./transform-2d";
import { Transform2D } from "./transform-2d";

export interface IConnectionSiteOptions {
    readonly x: number;
    readonly y: number;
    readonly angle?: number;
}

export interface IShapePropertiesOptions extends ITransform2DOptions {
    readonly geometry?: string;
    readonly fill?: FillOptions;
    readonly outline?: OutlineOptions;
    readonly effects?: IEffectsOptions;
    readonly connectionSites?: readonly IConnectionSiteOptions[];
}

/**
 * p:spPr — Shape properties (transform, geometry, fill, outline, effects).
 * Uses p: prefix in PresentationML context, though type is a:CT_ShapeProperties.
 */
export class ShapeProperties extends Xc {
    private readonly fillOptions?: FillOptions;

    public constructor(options: IShapePropertiesOptions) {
        super("p:spPr");

        this.fillOptions = options.fill;

        if (
            options.x !== undefined ||
            options.y !== undefined ||
            options.width !== undefined ||
            options.height !== undefined ||
            options.flipH !== undefined ||
            options.rotation !== undefined
        ) {
            this.root.push(new Transform2D(options));
        }

        if (options.geometry) {
            this.root.push(new PresetGeometry(options.geometry));
        } else {
            this.root.push(new PresetGeometry("rect"));
        }

        if (options.fill !== undefined) {
            this.root.push(buildFill(options.fill));
        } else {
            this.root.push(buildFill({ type: "none" }));
        }

        if (options.outline) {
            this.root.push(createOutlineCompat(options.outline));
        }

        if (options.effects) {
            this.root.push(new EffectList(options.effects));

            const scene3d = buildScene3D(options.effects);
            if (scene3d) this.root.push(scene3d);

            const shape3d = buildShape3D(options.effects);
            if (shape3d) this.root.push(shape3d);
        }

        if (options.connectionSites && options.connectionSites.length > 0) {
            const cxns = options.connectionSites.map((site) => {
                const attrs: Record<
                    string,
                    { readonly key: string; readonly value: string | number }
                > = {
                    pos: { key: "pos", value: `${site.x} ${site.y}` },
                };
                if (site.angle !== undefined) attrs.ang = { key: "ang", value: site.angle };
                return new BuilderElement({ name: "a:cxn", attributes: attrs });
            });
            this.root.push(new BuilderElement({ name: "a:cxnLst", children: cxns }));
        }
    }

    public override prepForXml(context: IContext<File>): IXmlableObject | undefined {
        const media = this.fillOptions ? extractBlipFillMedia(this.fillOptions) : undefined;
        if (media) {
            context.fileData?.Media.addImage(media.fileName, {
                data: media.data,
                fileName: media.fileName,
                type: media.type as "png",
                transformation: { pixels: { x: 0, y: 0 }, emus: { x: 0, y: 0 } },
            });
        }
        return super.prepForXml(context);
    }
}
