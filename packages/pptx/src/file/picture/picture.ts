import { BlipFill } from "@file/drawingml/blip-fill";
import { PresetGeometry } from "@file/drawingml/preset-geometry";
import { Transform2D } from "@file/drawingml/transform-2d";
import type { File } from "@file/file";
import type { IMediaData } from "@file/media/data";
import { BuilderElement, type Context, XmlComponent } from "@file/xml-components";
import { convertPixelsToEmu } from "@office-open/core";
import { emuPosition } from "@util/position";

import { PictureNonVisual } from "./picture-non-visual";

export interface PictureOptions {
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
 * Registers image with Media collection via toXml().
 * The ImageReplacer replaces `{fileName}` placeholder with actual rId.
 */
export class Picture extends XmlComponent {
  private static nextId = 100;
  private readonly imageData: IMediaData;

  public constructor(options: PictureOptions) {
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
          x: convertPixelsToEmu(options.width ?? 0),
          y: convertPixelsToEmu(options.height ?? 0),
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
            ...emuPosition(options),
          }),
          new PresetGeometry({ preset: "rect" }),
        ],
      }),
    );
  }

  /** Register image with the File's Media collection. */
  private registerMedia(context: Context): void {
    (context.fileData as File).media.addImage(this.imageData.fileName, this.imageData);
  }

  public override toXml(context: Context): string {
    this.registerMedia(context);
    return super.toXml(context);
  }
}
