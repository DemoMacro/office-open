import * as fs from "fs";
import { fileURLToPath } from "url";
import * as path from "path";

import { Presentation, Packer, Shape, SolidFill, Paragraph, Run, VideoFrame } from "../src";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const videoPath = path.join(__dirname, "assets/test-video.mp4");
const videoData = new Uint8Array(fs.readFileSync(videoPath));
const posterPath = path.join(__dirname, "assets/test-poster.png");
const posterData = new Uint8Array(fs.readFileSync(posterPath));

const pres = new Presentation({
    title: "Video Demo",
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
                                new Run({ text: "Video Embedding Demo", fontSize: 32, bold: true }),
                            ],
                        }),
                    ],
                }),
                new VideoFrame({
                    x: 50,
                    y: 100,
                    width: 480,
                    height: 270,
                    data: videoData,
                    type: "mp4",
                    name: "Big Buck Bunny",
                    poster: posterData,
                    posterType: "png",
                }),
                new Shape({
                    x: 50,
                    y: 390,
                    width: 500,
                    height: 40,
                    paragraphs: [
                        new Paragraph({
                            properties: { bulletNone: true },
                            children: [
                                new Run({
                                    text: "Video: Big Buck Bunny (360p, 10s, ~1MB MP4)",
                                    fontSize: 14,
                                    fill: new SolidFill("666666"),
                                }),
                            ],
                        }),
                    ],
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(pres);
fs.writeFileSync("My Presentation.pptx", buffer);
