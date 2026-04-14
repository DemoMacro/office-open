// Example of making content of section vertically aligned

import * as fs from "fs";

import { Document, Packer, Paragraph, Tab, TextRun, VerticalAlignSection } from "docx";

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
            properties: {
                verticalAlign: VerticalAlignSection.CENTER,
            },
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
