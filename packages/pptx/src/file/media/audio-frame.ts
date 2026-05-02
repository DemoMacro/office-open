import { PresetGeometry } from "@file/drawingml/preset-geometry";
import { Transform2D } from "@file/drawingml/transform-2d";
import type { File } from "@file/file";
import type { IMediaData } from "@file/media/data";
import { BuilderElement, type IContext, XmlComponent } from "@file/xml-components";
import { pixelsToEmus } from "@util/types";

const MEDIA_EXT_URI = "{CF1602FD-DB20-4165-A070-5F299619DA56}";
const P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main";

export type AudioType = "mp3" | "wav" | "wma" | "aac";

export interface IAudioFrameOptions {
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number;
    readonly data: Uint8Array;
    readonly type: AudioType;
    readonly name?: string;
}

/**
 * p:pic — An audio frame on a slide.
 *
 * Uses a media relationship for the audio file (via {media:fileName} placeholder).
 */
export class AudioFrame extends XmlComponent {
    private static nextId = 200;
    private readonly audioData: IMediaData;

    public constructor(options: IAudioFrameOptions) {
        super("p:pic");

        const id = AudioFrame.nextId++;
        const name = options.name ?? `Audio ${id}`;
        const mediaFileName = `${name.replace(/\s+/g, "_")}.${options.type}`;

        this.audioData = {
            type: options.type,
            fileName: mediaFileName,
            transformation: {
                pixels: { x: options.width ?? 0, y: options.height ?? 0 },
                emus: { x: pixelsToEmus(options.width ?? 0), y: pixelsToEmus(options.height ?? 0) },
            },
            data: options.data,
        };

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

        // p:blipFill — empty (no poster for audio, PowerPoint provides default icon)
        this.root.push(
            new BuilderElement({
                name: "p:blipFill",
                children: [
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
        (context.fileData as File)?.Media.addMedia(this.audioData.fileName, this.audioData);
        return super.prepForXml(context);
    }
}
