// Image blip effects (grayscale, luminance, duotone, tint, etc.)
import * as fs from "fs";

import { Document, ImageRun, Packer, Paragraph, TextRun } from "docx-plus";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            size: 32,
                            text: "Image Blip Effects",
                        }),
                    ],
                    spacing: { after: 400 },
                }),

                // 1. Grayscale
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "1. Grayscale",
                        }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("./demo/images/cat.jpg"),
                            blipEffects: { grayscale: true },
                            transformation: {
                                height: 150,
                                width: 150,
                            },
                            type: "jpg",
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // 2. Brightness + Contrast
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "2. Brightness (+30%) & Contrast (-20%)",
                        }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("./demo/images/cat.jpg"),
                            blipEffects: {
                                luminance: { bright: 30, contrast: -20 },
                            },
                            transformation: {
                                height: 150,
                                width: 150,
                            },
                            type: "jpg",
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // 3. Duotone
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "3. Duotone (Dark Blue + Light Gray)",
                        }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("./demo/images/cat.jpg"),
                            blipEffects: {
                                duotone: {
                                    color1: { value: "002060" },
                                    color2: { value: "D0CECE" },
                                },
                            },
                            transformation: {
                                height: 150,
                                width: 150,
                            },
                            type: "jpg",
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // 4. Tint
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "4. Tint (Warm, 40%)",
                        }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("./demo/images/cat.jpg"),
                            blipEffects: {
                                tint: { hue: 6000000, amt: 40 },
                            },
                            transformation: {
                                height: 150,
                                width: 150,
                            },
                            type: "jpg",
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // 5. BiLevel (Black & White threshold)
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "5. BiLevel (Threshold 50%)",
                        }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("./demo/images/cat.jpg"),
                            blipEffects: {
                                biLevel: { thresh: 50 },
                            },
                            transformation: {
                                height: 150,
                                width: 150,
                            },
                            type: "jpg",
                        }),
                    ],
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
