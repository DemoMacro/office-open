// Simple example to add text to a document

import * as fs from "fs";

import { CharacterSet, Document, Packer, Paragraph, Tab, TextRun } from "docx";

const font = fs.readFileSync("./demo/assets/Pacifico.ttf");

const doc = new Document({
    fonts: [{ characterSet: CharacterSet.ANSI, data: font, name: "Pacifico" }],
    sections: [
        {
            children: [
                new Paragraph({
                    run: {
                        font: "Pacifico",
                    },
                    children: [
                        new TextRun("Hello World"),
                        new TextRun({
                            text: "Foo Bar",
                            bold: true,
                            size: 40,
                            font: "Pacifico",
                        }),
                        new TextRun({
                            children: [new Tab(), "Github is the best"],
                            bold: true,
                            font: "Pacifico",
                        }),
                    ],
                }),
            ],
            properties: {},
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
