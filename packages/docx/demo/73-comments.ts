// Simple example to add comments to a document

import * as fs from "fs";

import {
    CommentRangeEnd,
    CommentRangeStart,
    CommentReference,
    Document,
    ImageRun,
    Packer,
    Paragraph,
    TextRun,
} from "docx-plus";

const doc = new Document({
    comments: {
        children: [
            {
                author: "Ray Chen",
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "some initial text content",
                            }),
                        ],
                    }),
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
                                text: "comment text content",
                            }),
                            new TextRun({ break: 1, text: "" }),
                            new TextRun({
                                bold: true,
                                text: "More text here",
                            }),
                        ],
                    }),
                ],
                date: new Date(),
                id: 0,
            },
            {
                author: "Bob Ross",
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "Some initial text content",
                            }),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "comment text content",
                            }),
                        ],
                    }),
                ],
                date: new Date(),
                id: 1,
            },
            {
                author: "John Doe",
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "Hello World",
                            }),
                        ],
                    }),
                ],
                date: new Date(),
                id: 2,
            },
            {
                author: "Beatriz",
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "Another reply",
                            }),
                        ],
                    }),
                ],
                date: new Date(),
                id: 3,
            },
        ],
    },
    sections: [
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun("Hello World"),
                        new CommentRangeStart(0),
                        new TextRun({
                            bold: true,
                            text: "Foo Bar",
                        }),
                        new CommentRangeEnd(0),
                        new TextRun({
                            bold: true,
                            children: [new CommentReference(0)],
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new CommentRangeStart(1),
                        new CommentRangeStart(2),
                        new CommentRangeStart(3),
                        new TextRun({
                            bold: true,
                            text: "Some text which need commenting",
                        }),
                        new CommentRangeEnd(1),
                        new TextRun({
                            bold: true,
                            children: [new CommentReference(1)],
                        }),
                        new CommentRangeEnd(2),
                        new TextRun({
                            bold: true,
                            children: [new CommentReference(2)],
                        }),
                        new CommentRangeEnd(3),
                        new TextRun({
                            bold: true,
                            children: [new CommentReference(3)],
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
