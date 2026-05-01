// Exporting the document as a stream

import * as fs from "fs";

import { Document, Packer, Paragraph, Tab, TextRun } from "docx-plus";

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

const stream = Packer.toStream(doc);
stream.pipe(fs.createWriteStream("My Document.docx"));
