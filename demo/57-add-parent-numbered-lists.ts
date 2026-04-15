// Numbered lists - Add parent number in sub number

import * as fs from "fs";

import {
    AlignmentType,
    Document,
    HeadingLevel,
    LevelFormat,
    Packer,
    Paragraph,
    convertInchesToTwip,
} from "docx-plus";

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
                                indent: { hanging: 260, left: convertInchesToTwip(0.5) },
                            },
                        },
                        text: "%1",
                    },
                    {
                        alignment: AlignmentType.START,
                        format: LevelFormat.DECIMAL,
                        level: 1,
                        style: {
                            paragraph: {
                                indent: {
                                    hanging: 1.25 * 260,
                                    left: 1.25 * convertInchesToTwip(0.5),
                                },
                            },
                            run: {
                                bold: true,
                                font: "Times New Roman",
                                size: 18,
                            },
                        },
                        text: "%1.%2",
                    },
                ],
                reference: "my-number-numbering-reference",
            },
        ],
    },
    sections: [
        {
            children: [
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    text: "How to make cake",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "my-number-numbering-reference",
                    },
                    text: "Step 1 - Add sugar",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "my-number-numbering-reference",
                    },
                    text: "Step 2 - Add wheat",
                }),
                new Paragraph({
                    numbering: {
                        level: 1,
                        reference: "my-number-numbering-reference",
                    },
                    text: "Step 2a - Stir the wheat in a circle",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "my-number-numbering-reference",
                    },
                    text: "Step 3 - Put in oven",
                }),
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    text: "How to make cake",
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
