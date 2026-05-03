// Simple example to add text to a document

import * as fs from "fs";

import { Document, Packer, Paragraph, Tab, TextRun } from "@office-open/docx";

const font = fs.readFileSync("./demo/assets/Pacifico.ttf");

const doc = new Document({
    fonts: [{ characterSet: "00", data: font, name: "Pacifico" }],
    sections: [
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun("Hello World"),
                        new TextRun({
                            text: "Foo Bar",
                            bold: true,
                            size: 40,
                        }),
                        new TextRun({
                            children: [new Tab(), "Github is the best"],
                            bold: true,
                        }),
                    ],
                }),
            ],
            properties: {},
        },
    ],
    styles: {
        default: {
            document: {
                run: {
                    font: "Pacifico",
                },
            },
        },
    },
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
