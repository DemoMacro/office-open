import { PresetGeometry } from "@file/drawingml/preset-geometry";
import { Transform2D } from "@file/drawingml/transform-2d";
import type { File } from "@file/file";
import type { IMediaData } from "@file/media/data";
import { BuilderElement, type IContext, XmlComponent } from "@file/xml-components";
import { pixelsToEmus } from "@util/types";

const MEDIA_EXT_URI = "{CF1602FD-DB20-4165-A070-5F299619DA56}";
const P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main";

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
}

/**
 * p:pic — A video frame on a slide.
 *
 * Uses two relationships:
 * - Image relationship for poster frame (via {fileName} placeholder)
 * - Media relationship for video file (via {media:fileName} placeholder)
 */
export class VideoFrame extends XmlComponent {
    private static nextId = 100;
    private readonly posterData?: IMediaData;
    private readonly videoData: IMediaData;

    public constructor(options: IVideoFrameOptions) {
        super("p:pic");

        const id = VideoFrame.nextId++;
        const name = options.name ?? `Video ${id}`;
        const mediaFileName = `${name.replace(/\s+/g, "_")}.${options.type}`;
        const posterType = options.posterType ?? "png";
        const posterFileName = `${name.replace(/\s+/g, "_")}_poster.${posterType}`;

        this.videoData = {
            type: options.type,
            fileName: mediaFileName,
            transformation: {
                pixels: { x: options.width ?? 0, y: options.height ?? 0 },
                emus: { x: pixelsToEmus(options.width ?? 0), y: pixelsToEmus(options.height ?? 0) },
            },
            data: options.data,
        };

        if (options.poster) {
            this.posterData = {
                type: posterType === "jpg" ? "jpg" : "png",
                fileName: posterFileName,
                transformation: {
                    pixels: { x: options.width ?? 0, y: options.height ?? 0 },
                    emus: { x: pixelsToEmus(options.width ?? 0), y: pixelsToEmus(options.height ?? 0) },
                },
                data: options.poster,
            };
        }

        // p:nvPicPr with p14:media extension
        this.root.push(
            new BuilderElement({
                name: "p:nvPicPr",
                children: [
                    new BuilderElement({
                        name: "a:cNvPr",
                        attributes: {
                            id: { key: "id", value: id },
                            name: { key: "name", value: name },
                            descr: { key: "descr", value: "" },
                        },
                    }),
                    new BuilderElement({
                        name: "a:cNvPicPr",
                        children: [
                            new BuilderElement({
                                name: "a:picLocks",
                                attributes: { noChangeAspect: { key: "noChangeAspect", value: "1" } },
                            }),
                        ],
                    }),
                    new BuilderElement({
                        name: "p:nvPr",
                        children: [
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
                                                    "r:embed": { key: "r:embed", value: `{media:${mediaFileName}}` },
                                                },
                                                namespaces: { p14: P14_NS },
                                            }),
                                        ],
                                    }),
                                ],
                            }),
                        ],
                    }),
                ],
            }),
        );

        // p:blipFill with poster image
        if (this.posterData) {
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
        }

        // p:spPr
        this.root.push(
            new BuilderElement({
                name: "p:spPr",
                children: [
                    new Transform2D({
                        x: pixelsToEmus(options.x ?? 0),
                        y: pixelsToEmus(options.y ?? 0),
                        width: pixelsToEmus(options.width ?? 0),
                        height: pixelsToEmus(options.height ?? 0),
                    }),
                    new PresetGeometry("rect"),
                ],
            }),
        );
    }

    public override prepForXml(context: IContext) {
        const file = context.fileData as File;
        if (this.posterData) {
            file?.Media.addImage(this.posterData.fileName, this.posterData);
        }
        file?.Media.addMedia(this.videoData.fileName, this.videoData);
        return super.prepForXml(context);
    }
}
