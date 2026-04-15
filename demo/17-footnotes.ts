// Footnotes

import * as fs from "fs";

import {
    AlignmentType,
    Document,
    FootnoteReferenceRun,
    ImageRun,
    LevelFormat,
    Packer,
    Paragraph,
    TextRun,
    convertInchesToTwip,
} from "docx-plus";

const doc = new Document({
    footnotes: {
        1: { children: [new Paragraph("Foo"), new Paragraph("Bar")] },
        2: {
            children: [
                new Paragraph("This footnote contains a numbered list:"),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "footnote-numbering",
                    },
                    text: "First item in the list",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "footnote-numbering",
                    },
                    text: "Second item in the list",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "footnote-numbering",
                    },
                    text: "Third item in the list",
                }),
            ],
        },
        3: {
            children: [
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("./demo/images/cat.jpg"),
                            transformation: {
                                height: 100,
                                width: 100,
                            },
                            type: "jpg",
                        }),
                        new TextRun({
                            text: "It's a cat",
                        }),
                    ],
                }),
            ],
        },
        4: { children: [new Paragraph("Foo1")] },
        5: { children: [new Paragraph("Test1")] },
        6: { children: [new Paragraph("My amazing reference1")] },
    },
    numbering: {
        config: [
            {
                levels: [
                    {
                        level: 0,
                        format: LevelFormat.DECIMAL,
                        text: "%1.",
                        alignment: AlignmentType.START,
                        style: {
                            paragraph: {
                                indent: {
                                    left: convertInchesToTwip(0.5),
                                    hanging: convertInchesToTwip(0.18),
                                },
                            },
                        },
                    },
                ],
                reference: "footnote-numbering",
            },
        ],
    },
    sections: [
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun({
                            children: ["Hello"],
                        }),
                        new FootnoteReferenceRun(1),
                        new TextRun({
                            children: [" World!"],
                        }),
                        new FootnoteReferenceRun(2),
                        new TextRun({
                            children: [" GitHub!"],
                        }),
                    ],
                }),
                new Paragraph({
                    children: [new TextRun("Hello World"), new FootnoteReferenceRun(3)],
                }),
            ],
        },
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun({
                            children: ["Hello"],
                        }),
                        new FootnoteReferenceRun(4),
                        new TextRun({
                            children: [" World!"],
                        }),
                        new FootnoteReferenceRun(5),
                    ],
                }),
                new Paragraph({
                    children: [new TextRun("Hello World Again"), new FootnoteReferenceRun(6)],
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
