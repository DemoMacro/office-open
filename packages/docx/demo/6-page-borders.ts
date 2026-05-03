// Example of how to change page borders

import * as fs from "fs";

import { Document, HeadingLevel, Packer, Paragraph, Tab, TextRun } from "@office-open/docx";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun("Hello World"),
                        new TextRun({
                            bold: true,
                            text: "Foo bar",
                        }),
                        new TextRun({
                            bold: true,
                            children: [new Tab(), "Github is the best"],
                        }),
                    ],
                }),
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    text: "Hello World",
                }),
                new Paragraph("Foo bar"),
                new Paragraph("Github is the best"),
            ],
            properties: {
                page: {
                    margin: {
                        bottom: 0,
                        left: 0,
                        right: 0,
                        top: 0,
                    },
                },
            },
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
