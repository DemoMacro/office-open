import type { File } from "@file/file";
import { BaseXmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";

import { EffectList, buildScene3D, buildShape3D, type IEffectsOptions } from "./effects";
import { buildFill, extractBlipFillMedia } from "./fill";
import type { FillOptions } from "./fill";
import { createOutlineCompat } from "./outline";
import type { OutlineOptions } from "./outline";
import { PresetGeometry } from "./preset-geometry";
import { Transform2D } from "./transform-2d";

export interface IConnectionSiteOptions {
    readonly x: number;
    readonly y: number;
    readonly angle?: number;
}

export interface IShapePropertiesOptions {
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number;
    readonly flipH?: boolean;
    readonly rotation?: number;
    readonly geometry?: string;
    readonly fill?: FillOptions;
    readonly outline?: OutlineOptions;
    readonly effects?: IEffectsOptions;
    readonly connectionSites?: readonly IConnectionSiteOptions[];
}

/**
 * p:spPr — Shape properties (transform, geometry, fill, outline, effects).
 * Lazy: stores options, builds XML object directly in prepForXml.
 */
export class ShapeProperties extends BaseXmlComponent {
    private readonly options: IShapePropertiesOptions;

    public constructor(options: IShapePropertiesOptions) {
        super("p:spPr");
        this.options = options;
    }

    public prepForXml(context: IContext<File>): IXmlableObject | undefined {
        const opts = this.options;
        const children: IXmlableObject[] = [];

        // Transform2D
        if (
            opts.x !== undefined ||
            opts.y !== undefined ||
            opts.width !== undefined ||
            opts.height !== undefined ||
            opts.flipH !== undefined ||
            opts.rotation !== undefined
        ) {
            const xfrmObj = new Transform2D(opts).prepForXml(context);
            if (xfrmObj) children.push(xfrmObj);
        }

        // PresetGeometry
        const geomObj = new PresetGeometry(opts.geometry ?? "rect").prepForXml(context);
        if (geomObj) children.push(geomObj);

        // Fill (register blipFill media — B-level side effect)
        const media = opts.fill ? extractBlipFillMedia(opts.fill) : undefined;
        if (media) {
            context.fileData?.Media.addImage(media.fileName, {
                data: media.data,
                fileName: media.fileName,
                type: media.type as "png",
                transformation: { pixels: { x: 0, y: 0 }, emus: { x: 0, y: 0 } },
            });
        }

        const fillComponent = buildFill(opts.fill !== undefined ? opts.fill : { type: "none" });
        const fillObj = fillComponent.prepForXml(context);
        if (fillObj) children.push(fillObj);

        // Outline
        if (opts.outline) {
            const outlineObj = createOutlineCompat(opts.outline).prepForXml(context);
            if (outlineObj) children.push(outlineObj);
        }

        // Effects
        if (opts.effects) {
            const effectObj = new EffectList(opts.effects).prepForXml(context);
            if (effectObj) children.push(effectObj);

            const scene3d = buildScene3D(opts.effects);
            if (scene3d) {
                const sceneObj = scene3d.prepForXml(context);
                if (sceneObj) children.push(sceneObj);
            }

            const shape3d = buildShape3D(opts.effects);
            if (shape3d) {
                const shapeObj = shape3d.prepForXml(context);
                if (shapeObj) children.push(shapeObj);
            }
        }

        // Connection sites
        if (opts.connectionSites && opts.connectionSites.length > 0) {
            const cxnChildren: IXmlableObject[] = [];
            for (const site of opts.connectionSites) {
                const siteAttrs: Record<string, string | number> = { pos: `${site.x} ${site.y}` };
                if (site.angle !== undefined) siteAttrs.ang = site.angle;
                cxnChildren.push({ "a:cxn": { _attr: siteAttrs } });
            }
            children.push({ "a:cxnLst": cxnChildren });
        }

        return { "p:spPr": children };
    }
}
