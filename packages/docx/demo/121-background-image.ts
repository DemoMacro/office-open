// Document background image: VML v:background + v:fill

import * as fs from "fs";

import { Document, Packer, Paragraph, TextRun } from "docx-plus";

const doc = new Document({
    background: {
        color: "E8E8E8",
        image: {
            data: fs.readFileSync("./demo/images/dog.png"),
            type: "png",
        },
    },
    sections: [
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "This document has a background image.",
                            bold: true,
                            size: 28,
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun(
                            "The image is rendered as a full-page background using VML v:fill.",
                        ),
                    ],
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
