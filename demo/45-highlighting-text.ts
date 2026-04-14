// Highlighting text

import * as fs from "fs";

import { AlignmentType, Document, Header, Packer, Paragraph, TextRun } from "docx-plus";

const doc = new Document({
    sections: [
        {
            children: [],
            headers: {
                default: new Header({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.RIGHT,
                            children: [
                                new TextRun({
                                    bold: true,
                                    color: "FF0000",
                                    font: {
                                        name: "Garamond",
                                    },
                                    highlight: "yellow",
                                    size: 24,
                                    text: "Hello World",
                                }),
                            ],
                        }),
                    ],
                }),
            },
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
