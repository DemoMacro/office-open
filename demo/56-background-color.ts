// Change background colour of whole document

import * as fs from "fs";

import { Document, Packer, Paragraph, Tab, TextRun } from "docx-plus";

const doc = new Document({
    background: {
        color: "C45911",
    },
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
