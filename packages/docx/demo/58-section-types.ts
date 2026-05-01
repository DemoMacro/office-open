// Usage of different Section Types

import * as fs from "fs";

import { Document, Packer, Paragraph, SectionType, TextRun } from "docx-plus";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun("Hello World"),
                        new TextRun({
                            bold: true,
                            text: "Foo Bar",
                        }),
                    ],
                }),
            ],
            properties: {},
        },
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun("Hello World"),
                        new TextRun({
                            bold: true,
                            text: "Foo Bar",
                        }),
                    ],
                }),
            ],
            properties: {
                type: SectionType.CONTINUOUS,
            },
        },
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun("Hello World"),
                        new TextRun({
                            bold: true,
                            text: "Foo Bar",
                        }),
                    ],
                }),
            ],
            properties: {
                type: SectionType.ODD_PAGE,
            },
        },
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun("Hello World"),
                        new TextRun({
                            bold: true,
                            text: "Foo Bar",
                        }),
                    ],
                }),
            ],
            properties: {
                type: SectionType.EVEN_PAGE,
            },
        },
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun("Hello World"),
                        new TextRun({
                            bold: true,
                            text: "Foo Bar",
                        }),
                    ],
                }),
            ],
            properties: {
                type: SectionType.NEXT_PAGE,
            },
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
