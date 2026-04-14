// Add text to header and footer

import * as fs from "fs";

import {
    AlignmentType,
    Document,
    Footer,
    Header,
    LevelFormat,
    Packer,
    Paragraph,
    convertInchesToTwip,
} from "docx";

const doc = new Document({
    numbering: {
        config: [
            {
                levels: [
                    {
                        alignment: AlignmentType.START,
                        format: LevelFormat.DECIMAL,
                        level: 0,
                        style: {
                            paragraph: {
                                indent: {
                                    hanging: convertInchesToTwip(0.18),
                                    left: convertInchesToTwip(0.5),
                                },
                            },
                        },
                        text: "%1.",
                    },
                ],
                reference: "footer-numbering",
            },
        ],
    },
    sections: [
        {
            children: [new Paragraph("Hello World")],
            footers: {
                default: new Footer({
                    children: [
                        new Paragraph("This footer contains a numbered list:"),
                        new Paragraph({
                            numbering: {
                                level: 0,
                                reference: "footer-numbering",
                            },
                            text: "First item in the list",
                        }),
                        new Paragraph({
                            numbering: {
                                level: 0,
                                reference: "footer-numbering",
                            },
                            text: "Second item in the list",
                        }),
                        new Paragraph({
                            numbering: {
                                level: 0,
                                reference: "footer-numbering",
                            },
                            text: "Third item in the list",
                        }),
                    ],
                }),
            },
            headers: {
                default: new Header({
                    children: [new Paragraph("Header text")],
                }),
            },
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
