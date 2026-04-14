// Japanese text - Need to use a Japanese font

import * as fs from "fs";

import { Document, HeadingLevel, Packer, Paragraph } from "docx-plus";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    text: "KFCを食べるのが好き",
                }),
                new Paragraph({
                    text: "こんにちは",
                }),
            ],
        },
    ],
    styles: {
        paragraphStyles: [
            {
                basedOn: "Normal",
                id: "Normal",
                name: "Normal",
                next: "Normal",
                quickFormat: true,
                run: {
                    font: "MS Gothic",
                },
            },
        ],
    },
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
