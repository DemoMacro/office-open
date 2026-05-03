// Numbered lists - With complex number text

import * as fs from "fs";

import { Document, LevelFormat, Packer, Paragraph } from "@office-open/docx";

const doc = new Document({
    numbering: {
        config: [
            {
                levels: [
                    {
                        format: LevelFormat.DECIMAL,
                        level: 0,
                        text: "%1",
                    },
                    {
                        format: LevelFormat.DECIMAL,
                        level: 1,
                        text: "%1.%2",
                    },
                    {
                        format: LevelFormat.DECIMAL,
                        level: 2,
                        text: "%1.%2.%3",
                    },
                ],
                reference: "ref1",
            },
        ],
    },
    sections: [
        {
            children: [
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "ref1",
                    },
                    text: "REF1 - lvl:0",
                }),
                new Paragraph({
                    numbering: {
                        level: 1,
                        reference: "ref1",
                    },
                    text: "REF1 - lvl:1",
                }),
                new Paragraph({
                    numbering: {
                        level: 2,
                        reference: "ref1",
                    },
                    text: "REF1  - lvl:2",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "ref1",
                    },
                    text: "REF1 - lvl:0",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "ref1",
                    },
                    text: "REF1 - lvl:0",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "ref1",
                    },
                    text: "REF1 - lvl:0",
                }),
                new Paragraph({
                    text: "Random text",
                }),
                new Paragraph({
                    numbering: {
                        instance: 1,
                        level: 0,
                        reference: "ref1",
                    },
                    text: "REF1 - inst:1 - lvl:0",
                }),
                new Paragraph({
                    numbering: {
                        instance: 0,
                        level: 0,
                        reference: "ref1",
                    },
                    text: "REF1 - inst:0 - lvl:0",
                }),
                new Paragraph({
                    numbering: {
                        instance: 0,
                        level: 0,
                        reference: "ref1",
                    },
                    text: "REF1 - inst:0 - lvl:0",
                }),
            ],
        },
    ],
});

// Used to export the file into a .docx file
const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
