// Image click and hover hyperlinks
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
                            text: "Image Hyperlink Demo",
                        }),
                    ],
                    spacing: { after: 400 },
                }),

                // 1. Image with click hyperlink
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "1. Image with Click Hyperlink",
                        }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [new TextRun("(Click the image below to open https://example.com)")],
                    spacing: { after: 100 },
                }),
                new Paragraph({
                    children: [
                        new ImageRun({
                            altText: {
                                description: "Click me!",
                                hyperlink: { click: "https://example.com" },
                                name: "link-image",
                                title: "Click to visit example.com",
                            },
                            data: fs.readFileSync("./demo/images/cat.jpg"),
                            transformation: {
                                height: 150,
                                width: 150,
                            },
                            type: "jpg",
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // 2. Image with both click and hover hyperlink
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "2. Image with Click + Hover Hyperlink",
                        }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new TextRun(
                            "(Click opens https://example.com, hover opens https://example.com/hover)",
                        ),
                    ],
                    spacing: { after: 100 },
                }),
                new Paragraph({
                    children: [
                        new ImageRun({
                            altText: {
                                description: "Click or hover me!",
                                hyperlink: {
                                    click: "https://example.com",
                                    hover: "https://example.com/hover",
                                },
                                name: "dual-link-image",
                                title: "Dual hyperlink image",
                            },
                            data: fs.readFileSync("./demo/images/cat.jpg"),
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
