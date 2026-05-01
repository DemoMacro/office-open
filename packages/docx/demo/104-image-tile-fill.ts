// Image tile fill mode (repeating image pattern)
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
                            text: "Image Tile Fill Demo",
                        }),
                    ],
                    spacing: { after: 400 },
                }),

                // 1. Default stretch (for comparison)
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "1. Default Stretch Fill (no tile)",
                        }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("./demo/images/cat.jpg"),
                            transformation: {
                                height: 100,
                                width: 300,
                            },
                            type: "jpg",
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // 2. Tile with 50% scale
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "2. Tile Fill (50% scale)",
                        }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("./demo/images/cat.jpg"),
                            tile: { sx: 50, sy: 50 },
                            transformation: {
                                height: 100,
                                width: 300,
                            },
                            type: "jpg",
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // 3. Tile with flip
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "3. Tile Fill (50% scale, XY flip)",
                        }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("./demo/images/cat.jpg"),
                            tile: { flip: "XY", sx: 50, sy: 50 },
                            transformation: {
                                height: 100,
                                width: 300,
                            },
                            type: "jpg",
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // 4. Tile with alignment
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "4. Tile Fill (50% scale, center aligned)",
                        }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("./demo/images/cat.jpg"),
                            tile: { align: "CENTER", sx: 50, sy: 50 },
                            transformation: {
                                height: 100,
                                width: 300,
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
