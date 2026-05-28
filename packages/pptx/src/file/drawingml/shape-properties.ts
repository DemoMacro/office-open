import type { File } from "@file/file";
import { BaseXmlComponent } from "@file/xml-components";
import type { Context, IXmlableObject } from "@file/xml-components";

import { createPptxEffectList, buildScene3D, buildShape3D, type EffectsOptions } from "./effects";
import { buildFill, extractBlipFillMedia } from "./fill";
import type { FillOptions } from "./fill";
import { createOutlineCompat } from "./outline";
import type { OutlineOptions } from "./outline";
import { PresetGeometry } from "./preset-geometry";
import { Transform2D } from "./transform-2d";

export interface ConnectionSiteOptions {
  readonly x: number;
  readonly y: number;
  readonly angle?: number;
}

export interface ShapePropertiesOptions {
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly flipHorizontal?: boolean;
  readonly rotation?: number;
  readonly geometry?: string;
  readonly fill?: FillOptions;
  readonly outline?: OutlineOptions;
  readonly effects?: EffectsOptions;
  readonly connectionSites?: readonly ConnectionSiteOptions[];
}

/**
 * p:spPr — Shape properties (transform, geometry, fill, outline, effects).
 * Lazy: stores options, builds XML object directly in prepForXml.
 */
export class ShapeProperties extends BaseXmlComponent {
  private readonly options: ShapePropertiesOptions;

  public constructor(options: ShapePropertiesOptions) {
    super("p:spPr");
    this.options = options;
  }

  public prepForXml(context: Context<File>): IXmlableObject | undefined {
    const opts = this.options;
    const children: IXmlableObject[] = [];

    // Transform2D
    if (
      opts.x !== undefined ||
      opts.y !== undefined ||
      opts.width !== undefined ||
      opts.height !== undefined ||
      opts.flipHorizontal !== undefined ||
      opts.rotation !== undefined
    ) {
      const xfrmObj = new Transform2D(opts).prepForXml(context);
      if (xfrmObj) children.push(xfrmObj);
    }

    // PresetGeometry
    const geomObj = new PresetGeometry({ preset: opts.geometry ?? "rect" }).prepForXml(context);
    if (geomObj) children.push(geomObj);

    // Fill (register blipFill media — B-level side effect)
    const media = opts.fill ? extractBlipFillMedia(opts.fill) : undefined;
    if (media) {
      context.fileData?.media.addImage(media.fileName, {
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
      const effectObj = createPptxEffectList(opts.effects);
      if (effectObj) {
        const effectXmlObj = effectObj.prepForXml(context);
        if (effectXmlObj) children.push(effectXmlObj);
      }

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
