import type { File } from "@file/file";
import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";

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
 */
export class ShapeProperties extends BaseXmlComponent {
  private readonly options: ShapePropertiesOptions;

  public constructor(options: ShapePropertiesOptions) {
    super("p:spPr");
    this.options = options;
  }

  public override toXml(context: Context<File>): string {
    const opts = this.options;
    const parts: string[] = [];

    // Transform2D
    if (
      opts.x !== undefined ||
      opts.y !== undefined ||
      opts.width !== undefined ||
      opts.height !== undefined ||
      opts.flipHorizontal !== undefined ||
      opts.rotation !== undefined
    ) {
      parts.push(new Transform2D(opts).toXml(context));
    }

    // PresetGeometry
    parts.push(new PresetGeometry({ preset: opts.geometry ?? "rect" }).toXml(context));

    // Fill (register blipFill media — side effect)
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
    parts.push(fillComponent.toXml(context));

    // Outline
    if (opts.outline) {
      parts.push(createOutlineCompat(opts.outline).toXml(context));
    }

    // Effects
    if (opts.effects) {
      const effectObj = createPptxEffectList(opts.effects);
      if (effectObj) parts.push(effectObj.toXml(context));

      const scene3d = buildScene3D(opts.effects);
      if (scene3d) parts.push(scene3d.toXml(context));

      const shape3d = buildShape3D(opts.effects);
      if (shape3d) parts.push(shape3d.toXml(context));
    }

    // Connection sites
    if (opts.connectionSites && opts.connectionSites.length > 0) {
      const cxnParts: string[] = [];
      for (const site of opts.connectionSites) {
        const ang = site.angle !== undefined ? ` ang="${site.angle}"` : "";
        cxnParts.push(`<a:cxn pos="${site.x} ${site.y}"${ang}/>`);
      }
      parts.push(`<a:cxnLst>${cxnParts.join("")}</a:cxnLst>`);
    }

    return `<p:spPr>${parts.join("")}</p:spPr>`;
  }
}
