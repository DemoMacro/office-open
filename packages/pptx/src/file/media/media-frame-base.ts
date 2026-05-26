import type { AnimationOptions } from "@file/animation/types";
/**
 * Shared base class for media frames (video, audio) on slides.
 * @module
 */
import { PresetGeometry } from "@file/drawingml/preset-geometry";
import { Transform2D } from "@file/drawingml/transform-2d";
import type { File } from "@file/file";
import type { IMediaData } from "@file/media/data";
import { BuilderElement, type Context, XmlComponent } from "@file/xml-components";
import { convertPixelsToEmu } from "@office-open/core";
import { emuPosition } from "@util/position";

/**
 * Common options for all media frames.
 * @internal
 */
export interface MediaFrameBaseOptions {
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly data: Uint8Array;
  readonly type: IMediaData["type"];
  readonly name?: string;
  readonly animation?: AnimationOptions;
}

/**
 * Builds an IMediaData object from pixel dimensions and raw data.
 * Generic parameter preserves the specific type literal for correct union narrowing.
 */
function buildMediaData<T extends IMediaData["type"]>(
  type: T,
  fileName: string,
  width: number,
  height: number,
  data: Uint8Array,
): IMediaData & { readonly type: T } {
  return {
    type,
    fileName,
    transformation: {
      pixels: { x: width, y: height },
      emus: { x: convertPixelsToEmu(width), y: convertPixelsToEmu(height) },
    },
    data,
  } as IMediaData & { readonly type: T };
}

/**
 * Base class for media frames (video, audio) on slides.
 *
 * Encapsulates the common structure: ID assignment, media data building,
 * p:spPr (Transform2D + PresetGeometry), p:blipFill, p:nvPicPr parts,
 * and media registration in prepForXml.
 */
export abstract class MediaFrameBase extends XmlComponent {
  private readonly shapeId: number;
  private readonly animationOptions?: AnimationOptions;
  protected readonly mediaData: IMediaData;
  protected readonly posterData?: IMediaData;

  protected constructor(
    options: MediaFrameBaseOptions,
    id: number,
    mediaFileName: string,
    params: {
      readonly extUri: string;
      readonly cNvPrPrefix: string;
      readonly nvPrExtraChildren?: BuilderElement[];
      readonly posterBytes?: Uint8Array;
      readonly posterType?: string;
      readonly posterFileName?: string;
    },
  ) {
    super("p:pic");

    this.shapeId = id;
    this.animationOptions = options.animation;

    const name = options.name ?? `Media ${id}`;
    const w = options.width ?? 0;
    const h = options.height ?? 0;

    this.mediaData = buildMediaData(options.type, mediaFileName, w, h, options.data);

    // Poster data (video only)
    if (params.posterBytes) {
      const pt = params.posterType ?? "png";
      const pfn = params.posterFileName ?? `${name.replace(/\s+/g, "_")}_poster.${pt}`;
      this.posterData = buildMediaData(pt === "jpg" ? "jpg" : "png", pfn, w, h, params.posterBytes);
    }

    // p:nvPicPr
    const nvPrChildren: BuilderElement[] = [...(params.nvPrExtraChildren ?? [])];
    nvPrChildren.push(
      new BuilderElement({
        name: "p:extLst",
        children: [
          new BuilderElement({
            name: "p:ext",
            attributes: { uri: { key: "uri", value: params.extUri } },
            children: [
              new BuilderElement({
                name: "p14:media",
                attributes: {
                  "r:embed": {
                    key: "r:embed",
                    value: `{media:${mediaFileName}}`,
                  },
                  "xmlns:p14": {
                    key: "xmlns:p14",
                    value: "http://schemas.microsoft.com/office/powerpoint/2010/main",
                  },
                },
              }),
            ],
          }),
        ],
      }),
    );

    this.root.push(
      new BuilderElement({
        name: "p:nvPicPr",
        children: [
          new BuilderElement({
            name: `${params.cNvPrPrefix}:cNvPr`,
            attributes: {
              id: { key: "id", value: id },
              name: { key: "name", value: name },
              descr: { key: "descr", value: "" },
            },
          }),
          new BuilderElement({
            name: `${params.cNvPrPrefix}:cNvPicPr`,
            children: [
              new BuilderElement({
                name: "a:picLocks",
                attributes: {
                  noChangeAspect: { key: "noChangeAspect", value: "1" },
                },
              }),
            ],
          }),
          new BuilderElement({
            name: "p:nvPr",
            children: nvPrChildren,
          }),
        ],
      }),
    );

    // p:blipFill
    const blipFillChildren: BuilderElement[] = [];
    if (this.posterData) {
      blipFillChildren.push(
        new BuilderElement({
          name: "a:blip",
          attributes: {
            "r:embed": { key: "r:embed", value: `{${this.posterData.fileName}}` },
          },
        }),
      );
    }
    blipFillChildren.push(
      new BuilderElement({
        name: "a:stretch",
        children: [new BuilderElement({ name: "a:fillRect" })],
      }),
    );
    this.root.push(new BuilderElement({ name: "p:blipFill", children: blipFillChildren }));

    // p:spPr
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

  public get ShapeId(): number {
    return this.shapeId;
  }

  public get Animation(): AnimationOptions | undefined {
    return this.animationOptions;
  }

  public override prepForXml(context: Context) {
    const file = context.fileData as File;
    if (this.posterData) {
      file?.Media.addImage(this.posterData.fileName, this.posterData);
    }
    file?.Media.addMedia(this.mediaData.fileName, this.mediaData);
    return super.prepForXml(context);
  }
}
