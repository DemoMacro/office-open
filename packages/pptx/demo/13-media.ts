import * as fs from "fs";

import { Presentation, Packer, Shape, SolidFill, Paragraph, Run, VideoFrame, AudioFrame } from "../src";

// Minimal MP4 file (1x1 pixel, 0.1s) for testing
const minimalMp4 = new Uint8Array([
    0x00, 0x00, 0x00, 0x20, 0x66, 0x74, 0x79, 0x70, 0x69, 0x73, 0x6F, 0x6D,
    0x00, 0x00, 0x02, 0x00, 0x69, 0x73, 0x6F, 0x6D, 0x69, 0x73, 0x6F, 0x32,
    0x61, 0x76, 0x63, 0x31, 0x6D, 0x70, 0x34, 0x31, 0x00, 0x00, 0x00, 0x00,
]);

// Minimal WAV file (44-byte header + silence) for testing
const minimalWav = new Uint8Array(44);
minimalWav.set([0x52, 0x49, 0x46, 0x46], 0); // "RIFF"
minimalWav.set([0x57, 0x41, 0x56, 0x45], 8); // "WAVE"
minimalWav.set([0x66, 0x6D, 0x74, 0x20], 12); // "fmt "

const pres = new Presentation({
    title: "Media Demo",
    creator: "Demo",
    slides: [
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 500,
                    height: 50,
                    paragraphs: [
                        new Paragraph({
                            properties: { alignment: "ctr", bulletNone: true },
                            children: [
                                new Run({ text: "Video & Audio Demo", fontSize: 32, bold: true }),
                            ],
                        }),
                    ],
                }),
                new Shape({
                    x: 50,
                    y: 90,
                    width: 500,
                    height: 30,
                    paragraphs: [
                        new Paragraph({
                            properties: { bulletNone: true },
                            children: [
                                new Run({
                                    text: "Video frame (MP4) — placeholder below",
                                    fontSize: 16,
                                    fill: new SolidFill("666666"),
                                }),
                            ],
                        }),
                    ],
                }),
                new VideoFrame({
                    x: 50,
                    y: 130,
                    width: 480,
                    height: 270,
                    data: minimalMp4,
                    type: "mp4",
                    name: "Sample Video",
                }),
                new Shape({
                    x: 50,
                    y: 420,
                    width: 500,
                    height: 30,
                    paragraphs: [
                        new Paragraph({
                            properties: { bulletNone: true },
                            children: [
                                new Run({
                                    text: "Audio frame (WAV) — placeholder below",
                                    fontSize: 16,
                                    fill: new SolidFill("666666"),
                                }),
                            ],
                        }),
                    ],
                }),
                new AudioFrame({
                    x: 50,
                    y: 460,
                    width: 48,
                    height: 48,
                    data: minimalWav,
                    type: "wav",
                    name: "Sample Audio",
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(pres);
fs.writeFileSync("My Presentation.pptx", buffer);
