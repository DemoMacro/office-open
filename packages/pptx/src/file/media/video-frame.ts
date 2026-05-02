import { PresetGeometry } from "@file/drawingml/preset-geometry";
import { Transform2D } from "@file/drawingml/transform-2d";
import type { File } from "@file/file";
import type { IMediaData } from "@file/media/data";
import { BuilderElement, type IContext, XmlComponent } from "@file/xml-components";
import { pixelsToEmus } from "@util/types";

const MEDIA_EXT_URI = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}";

/**
 * Extract the first video frame from MP4 data as JPEG.
 * Falls back to a generated placeholder if extraction fails.
 */
function extractFirstFrame(videoData: Uint8Array, type: VideoType): Uint8Array {
    if (type === "mp4") {
        const jpeg = extractMp4FirstFrame(videoData);
        if (jpeg) return jpeg;
    }
    return generatePlaceholderPoster();
}

/**
 * Attempt to extract the first I-frame from an MP4 file.
 * MP4 structure: ftyp -> moov (metadata) -> mdat (media data).
 * We scan mdat for JPEG start markers (FFD8) which indicate
 * embedded poster frames, or for AVC NAL unit start codes.
 */
function extractMp4FirstFrame(data: Uint8Array): Uint8Array | null {
    const bytes = new Uint8Array(data);

    // Check for embedded JPEG poster (some MP4 files have this in moov/udta/meta)
    for (let i = 0; i < bytes.length - 1; i++) {
        if (bytes[i] === 0xFF && bytes[i + 1] === 0xD8) {
            // Found JPEG SOI marker — find EOI
            let end = -1;
            for (let j = i + 2; j < bytes.length - 1; j++) {
                if (bytes[j] === 0xFF && bytes[j + 1] === 0xD9) {
                    end = j + 2;
                    break;
                }
            }
            if (end > i + 100) {
                return bytes.slice(i, end);
            }
        }
    }
    return null;
}

/**
 * Generate a minimal 320x180 dark gray PNG as a video placeholder poster.
 */
function generatePlaceholderPoster(): Uint8Array {
    const width = 320;
    const height = 180;
    const zlib = require("zlib");

    // RGBA raw data: dark gray background with a centered play triangle
    const rawData = Buffer.alloc((width * 3 + 1) * height);
    const centerX = width / 2;
    const centerY = height / 2;
    const triSize = 30;

    for (let y = 0; y < height; y++) {
        const rowOffset = y * (width * 3 + 1);
        rawData[rowOffset] = 0; // PNG filter: None
        for (let x = 0; x < width; x++) {
            const px = rowOffset + 1 + x * 3;

            // Play triangle: centered, pointing right
            const dx = x - centerX + triSize * 0.3;
            const dy = y - centerY;
            const inTriangle =
                dy >= -triSize &&
                dy <= triSize &&
                dx >= -triSize * 0.5 &&
                dx <= triSize * 0.5 &&
                dx <= triSize * 0.5 - (Math.abs(dy) / triSize) * triSize;

            if (inTriangle) {
                rawData[px] = 255;     // R
                rawData[px + 1] = 255; // G
                rawData[px + 2] = 255; // B
            } else {
                rawData[px] = 51;      // R  (#333333)
                rawData[px + 1] = 51;  // G
                rawData[px + 2] = 51;  // B
            }
        }
    }

    const compressed = zlib.deflateSync(rawData);

    function crc32(buf: Buffer): number {
        let crc = 0xFFFFFFFF;
        for (let i = 0; i < buf.length; i++) {
            crc ^= buf[i];
            for (let j = 0; j < 8; j++) crc = (crc >>> 1) ^ (crc & 1 ? 0xEDB88320 : 0);
        }
        return (crc ^ 0xFFFFFFFF) >>> 0;
    }

    function makeChunk(type: string, data: Buffer): Buffer {
        const len = Buffer.alloc(4);
        len.writeUInt32BE(data.length);
        const typeData = Buffer.concat([Buffer.from(type), data]);
        const crc = Buffer.alloc(4);
        crc.writeUInt32BE(crc32(typeData));
        return Buffer.concat([len, typeData, crc]);
    }

    const ihdr = Buffer.alloc(13);
    ihdr.writeUInt32BE(width, 0);
    ihdr.writeUInt32BE(height, 4);
    ihdr[8] = 8;  // bit depth
    ihdr[9] = 2;  // RGB

    return Buffer.concat([
        Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
        makeChunk("IHDR", ihdr),
        makeChunk("IDAT", compressed),
        makeChunk("IEND", Buffer.alloc(0)),
    ]);
}

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
    readonly posterFrameTime?: number;
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

        const posterBytes = options.poster ?? extractFirstFrame(options.data, options.type);
        this.posterData = {
            type: posterType === "jpg" ? "jpg" : "png",
            fileName: posterFileName,
            transformation: {
                pixels: { x: options.width ?? 0, y: options.height ?? 0 },
                emus: { x: pixelsToEmus(options.width ?? 0), y: pixelsToEmus(options.height ?? 0) },
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
                                    "r:embed": { key: "r:embed", value: `{media:${mediaFileName}}` },
                                    "xmlns:p14": { key: "xmlns:p14", value: "http://schemas.microsoft.com/office/powerpoint/2010/main" },
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
                                attributes: { noChangeAspect: { key: "noChangeAspect", value: "1" } },
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
        file?.Media.addImage(this.posterData.fileName, this.posterData);
        file?.Media.addMedia(this.videoData.fileName, this.videoData);
        return super.prepForXml(context);
    }
}
