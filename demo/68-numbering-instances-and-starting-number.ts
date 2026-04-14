import * as fs from "fs";

import { Document, LevelFormat, Packer, Paragraph } from "docx-plus";

const doc = new Document({
    numbering: {
        config: [
            {
                levels: [
                    {
                        format: LevelFormat.DECIMAL,
                        level: 0,
                        start: 10,
                        text: "%1",
                    },
                ],
                reference: "ref1",
            },
            {
                levels: [
                    {
                        format: LevelFormat.DECIMAL,
                        level: 0,
                        text: "%1",
                    },
                ],
                reference: "ref2",
            },
        ],
    },
    sections: [
        {
            children: [
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
                        instance: 1,
                        level: 0,
                        reference: "ref1",
                    },
                    text: "REF1 - inst:1 - lvl:0",
                }),
                new Paragraph({
                    numbering: {
                        instance: 1,
                        level: 0,
                        reference: "ref2",
                    },
                    text: "REF2 - inst:0 - lvl:0",
                }),
                new Paragraph({
                    numbering: {
                        instance: 1,
                        level: 0,
                        reference: "ref2",
                    },
                    text: "REF2 - inst:0 - lvl:0",
                }),
            ],
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
