// Move + offset header and footer

import * as fs from "fs";

import { Document, Footer, Header, Packer, PageBreak, Paragraph, TextRun } from "@office-open/docx";

const doc = new Document({
    evenAndOddHeaderAndFooters: true,
    sections: [
        {
            children: [
                new Paragraph({
                    children: [new TextRun("Hello World 1"), new PageBreak()],
                }),
                new Paragraph({
                    children: [new TextRun("Hello World 2"), new PageBreak()],
                }),
                new Paragraph({
                    children: [new TextRun("Hello World 3"), new PageBreak()],
                }),
                new Paragraph({
                    children: [new TextRun("Hello World 4"), new PageBreak()],
                }),
                new Paragraph({
                    children: [new TextRun("Hello World 5"), new PageBreak()],
                }),
            ],
            footers: {
                default: new Footer({
                    children: [
                        new Paragraph({
                            text: "Odd Footer text",
                        }),
                    ],
                }),
                even: new Footer({
                    children: [
                        new Paragraph({
                            text: "Even Cool Footer text",
                        }),
                    ],
                }),
            },
            headers: {
                default: new Header({
                    children: [
                        new Paragraph({
                            text: "Odd Header text",
                        }),
                        new Paragraph({
                            text: "Odd - Some more header text",
                        }),
                    ],
                }),
                even: new Header({
                    children: [
                        new Paragraph({
                            text: "Even header text",
                        }),
                        new Paragraph({
                            text: "Even - Some more header text",
                        }),
                    ],
                }),
            },
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
