// Demo: SmartArt - tree-shaped diagrams with nested nodes
import * as fs from "fs";

import { Document, Packer, Paragraph, TextRun, SmartArtRun } from "docx-plus";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "SmartArt Demo",
                            bold: true,
                            size: 32,
                        }),
                    ],
                    spacing: { after: 400 },
                }),

                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "Process Diagram",
                            size: 28,
                        }),
                    ],
                    spacing: { after: 200 },
                }),

                new Paragraph({
                    children: [
                        new SmartArtRun({
                            data: {
                                nodes: [
                                    {
                                        text: "Main Idea",
                                        children: [{ text: "Detail A" }, { text: "Detail B" }],
                                    },
                                    { text: "Support" },
                                ],
                            },
                            transformation: { width: 500, height: 300 },
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                new Paragraph({
                    children: [
                        new TextRun({
                            text: "Note: Open this document in Microsoft Word to see the SmartArt rendered.",
                            italics: true,
                            color: "888888",
                        }),
                    ],
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
