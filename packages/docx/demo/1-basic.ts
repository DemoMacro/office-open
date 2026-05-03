// Simple example to add text to a document

import * as fs from "fs";

import { Document, Packer, Paragraph, Tab, TextRun } from "@office-open/docx";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun("Hello World"),
                        new TextRun({
                            bold: true,
                            size: 40,
                            text: "Foo Bar",
                        }),
                        new TextRun({
                            bold: true,
                            children: [new Tab(), "Github is the best"],
                        }),
                    ],
                }),
            ],
            properties: {},
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
