// Creates two paragraphs, one with a border and one without

import * as fs from "fs";

import { BorderStyle, Document, Packer, Paragraph, TextRun } from "docx-plus";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph("No border!"),
                new Paragraph({
                    border: {
                        bottom: {
                            color: "auto",
                            size: 6,
                            space: 1,
                            style: BorderStyle.SINGLE,
                        },
                        top: {
                            color: "auto",
                            size: 6,
                            space: 1,
                            style: BorderStyle.SINGLE,
                        },
                    },
                    text: "I have borders on my top and bottom sides!",
                }),
                new Paragraph({
                    border: {
                        top: {
                            color: "auto",
                            size: 6,
                            space: 1,
                            style: BorderStyle.SINGLE,
                        },
                    },
                    text: "",
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "This will ",
                        }),
                        new TextRun({
                            border: {
                                color: "auto",
                                size: 6,
                                space: 1,
                                style: BorderStyle.SINGLE,
                            },
                            text: "have a border.",
                        }),
                        new TextRun({
                            text: " This will not.",
                        }),
                    ],
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
