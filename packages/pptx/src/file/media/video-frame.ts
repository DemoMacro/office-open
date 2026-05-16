import type { IAnimationOptions } from "@file/animation/types";
import { PresetGeometry } from "@file/drawingml/preset-geometry";
import { Transform2D } from "@file/drawingml/transform-2d";
import type { File } from "@file/file";
import type { IMediaData } from "@file/media/data";
import { BuilderElement, type IContext, XmlComponent } from "@file/xml-components";
import { convertPixelsToEmu } from "@office-open/core";

const MEDIA_EXT_URI = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}";

/** Minimal 1x1 transparent PNG (67 bytes). */
const MINIMAL_PNG = new Uint8Array([
    137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0,
    0, 0, 31, 21, 196, 137, 0, 0, 0, 13, 73, 68, 65, 84, 8, 215, 99, 24, 5, 163, 0, 0, 0, 2, 0, 1,
    226, 33, 188, 51, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
]);

export type VideoType = "mp4" | "mov" | "wmv" | "avi";
export type PosterType = "png" | "jpg";

export interface IVideoFrameOptions {
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number;
    readonly data: Uint8Array;
    readonly type: VideoType;
    readonly name?: string;
    readonly poster?: Uint8Array;
    readonly posterType?: PosterType;
    readonly animation?: IAnimationOptions;
}

/**
 * p:pic — A video frame on a slide.
 *
 * Uses three relationships:
 * - Image relationship for poster frame (via {posterFileName} placeholder)
 * - Video relationship for video file (via {video:fileName} placeholder in a:videoFile r:link)
 * - Media relationship for video file (via {media:fileName} placeholder in p14:media r:embed)
 */
export class VideoFrame extends XmlComponent {
    private static nextId = 100;
    private readonly posterData: IMediaData;
    private readonly videoData: IMediaData;
    private readonly shapeId: number;
    private readonly animationOptions?: IAnimationOptions;

    public constructor(options: IVideoFrameOptions) {
        super("p:pic");

        const id = VideoFrame.nextId++;
        this.shapeId = id;
        this.animationOptions = options.animation;
        const name = options.name ?? `Video ${id}`;
        const mediaFileName = `${name.replace(/\s+/g, "_")}.${options.type}`;
        const posterBytes = options.poster ?? MINIMAL_PNG;
        const posterType = options.posterType ?? "png";
        const posterFileName = `${name.replace(/\s+/g, "_")}_poster.${posterType}`;

        this.videoData = {
            type: options.type,
            fileName: mediaFileName,
            transformation: {
                pixels: { x: options.width ?? 0, y: options.height ?? 0 },
                emus: {
                    x: convertPixelsToEmu(options.width ?? 0),
                    y: convertPixelsToEmu(options.height ?? 0),
                },
            },
            data: options.data,
        };

        this.posterData = {
            type: posterType === "jpg" ? "jpg" : "png",
            fileName: posterFileName,
            transformation: {
                pixels: { x: options.width ?? 0, y: options.height ?? 0 },
                emus: {
                    x: convertPixelsToEmu(options.width ?? 0),
                    y: convertPixelsToEmu(options.height ?? 0),
                },
            },
            data: posterBytes,
        };

        // p:nvPicPr with a:videoFile + p14:media
        const nvPrChildren: BuilderElement[] = [
            new BuilderElement({
                name: "a:videoFile",
                attributes: {
                    "r:link": { key: "r:link", value: `{video:${mediaFileName}}` },
                },
            }),
            new BuilderElement({
                name: "p:extLst",
                children: [
                    new BuilderElement({
                        name: "p:ext",
                        attributes: { uri: { key: "uri", value: MEDIA_EXT_URI } },
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
        ];

        this.root.push(
            new BuilderElement({
                name: "p:nvPicPr",
                children: [
                    new BuilderElement({
                        name: "p:cNvPr",
                        attributes: {
                            id: { key: "id", value: id },
                            name: { key: "name", value: name },
                            descr: { key: "descr", value: "" },
                        },
                    }),
                    new BuilderElement({
                        name: "p:cNvPicPr",
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

        // p:blipFill with poster image (required by PowerPoint)
        this.root.push(
            new BuilderElement({
                name: "p:blipFill",
                children: [
                    new BuilderElement({
                        name: "a:blip",
                        attributes: {
                            "r:embed": { key: "r:embed", value: `{${posterFileName}}` },
                        },
                    }),
                    new BuilderElement({
                        name: "a:stretch",
                        children: [new BuilderElement({ name: "a:fillRect" })],
                    }),
                ],
            }),
        );

        // p:spPr
        this.root.push(
            new BuilderElement({
                name: "p:spPr",
                children: [
                    new Transform2D({
                        x: convertPixelsToEmu(options.x ?? 0),
                        y: convertPixelsToEmu(options.y ?? 0),
                        width: convertPixelsToEmu(options.width ?? 0),
                        height: convertPixelsToEmu(options.height ?? 0),
                    }),
                    new PresetGeometry({ preset: "rect" }),
                ],
            }),
        );
    }

    public get ShapeId(): number {
        return this.shapeId;
    }

    public get Animation(): IAnimationOptions | undefined {
        return this.animationOptions;
    }

    public override prepForXml(context: IContext) {
        const file = context.fileData as File;
        file?.Media.addImage(this.posterData.fileName, this.posterData);
        file?.Media.addMedia(this.videoData.fileName, this.videoData);
        return super.prepForXml(context);
    }
}
