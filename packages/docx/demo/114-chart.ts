// Demo: Chart - bar/column chart with multiple series
import * as fs from "fs";

import { Document, Packer, Paragraph, TextRun, ChartRun } from "@office-open/docx";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "Chart Demo",
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
                            text: "Bar Chart",
                            size: 28,
                        }),
                    ],
                    spacing: { after: 200 },
                }),

                new Paragraph({
                    children: [
                        new ChartRun({
                            type: "column",
                            data: {
                                categories: ["Q1", "Q2", "Q3", "Q4"],
                                series: [
                                    { name: "2024", values: [120, 150, 180, 200] },
                                    { name: "2025", values: [140, 170, 210, 250] },
                                ],
                            },
                            title: "Quarterly Revenue",
                            transformation: { width: 500, height: 300 },
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                new Paragraph({
                    children: [
                        new TextRun({
                            text: "Note: Open this document in Microsoft Word to see the chart rendered.",
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
