// Numbered lists
// The lists can also be restarted by specifying the instance number

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
                        format: LevelFormat.UPPER_ROMAN,
                        level: 0,
                        style: {
                            paragraph: {
                                indent: {
                                    hanging: convertInchesToTwip(0.18),
                                    left: convertInchesToTwip(0.5),
                                },
                            },
                        },
                        text: "%1",
                    },
                ],
                reference: "my-crazy-reference",
            },
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
                        text: "%1",
                    },
                ],
                reference: "my-number-numbering-reference",
            },
            {
                levels: [
                    {
                        alignment: AlignmentType.START,
                        format: LevelFormat.DECIMAL_ZERO,
                        level: 0,
                        style: {
                            paragraph: {
                                indent: {
                                    hanging: convertInchesToTwip(0.18),
                                    left: convertInchesToTwip(0.5),
                                },
                            },
                        },
                        text: "[%1]",
                    },
                ],
                reference: "padded-numbering-reference",
            },
        ],
    },
    sections: [
        {
            children: [
                new Paragraph({
                    contextualSpacing: true,
                    numbering: {
                        level: 0,
                        reference: "my-crazy-reference",
                    },
                    spacing: {
                        before: 200,
                    },
                    text: "line with contextual spacing",
                }),
                new Paragraph({
                    contextualSpacing: true,
                    numbering: {
                        level: 0,
                        reference: "my-crazy-reference",
                    },
                    spacing: {
                        before: 200,
                    },
                    text: "line with contextual spacing",
                }),
                new Paragraph({
                    contextualSpacing: false,
                    numbering: {
                        level: 0,
                        reference: "my-crazy-reference",
                    },
                    spacing: {
                        before: 200,
                    },
                    text: "line without contextual spacing",
                }),
                new Paragraph({
                    contextualSpacing: false,
                    numbering: {
                        level: 0,
                        reference: "my-crazy-reference",
                    },
                    spacing: {
                        before: 200,
                    },
                    text: "line without contextual spacing",
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
                        level: 0,
                        reference: "my-number-numbering-reference",
                    },
                    text: "Step 3 - Put in oven",
                }),
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    text: "Next",
                }),
                new Paragraph({
                    numbering: {
                        instance: 2,
                        level: 0,
                        reference: "padded-numbering-reference",
                    },
                    text: "test",
                }),
                new Paragraph({
                    numbering: {
                        instance: 2,
                        level: 0,
                        reference: "padded-numbering-reference",
                    },
                    text: "test",
                }),
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    text: "Next",
                }),
                new Paragraph({
                    numbering: {
                        instance: 3,
                        level: 0,
                        reference: "padded-numbering-reference",
                    },
                    text: "test",
                }),
                new Paragraph({
                    numbering: {
                        instance: 3,
                        level: 0,
                        reference: "padded-numbering-reference",
                    },
                    text: "test",
                }),
                new Paragraph({
                    numbering: {
                        instance: 3,
                        level: 0,
                        reference: "padded-numbering-reference",
                    },
                    text: "test",
                }),
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    text: "Next",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "padded-numbering-reference",
                    },
                    text: "test",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "padded-numbering-reference",
                    },
                    text: "test",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "padded-numbering-reference",
                    },
                    text: "test",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "padded-numbering-reference",
                    },
                    text: "test",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "padded-numbering-reference",
                    },
                    text: "test",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "padded-numbering-reference",
                    },
                    text: "test",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "padded-numbering-reference",
                    },
                    text: "test",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "padded-numbering-reference",
                    },
                    text: "test",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "padded-numbering-reference",
                    },
                    text: "test",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "padded-numbering-reference",
                    },
                    text: "test",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "padded-numbering-reference",
                    },
                    text: "test",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "padded-numbering-reference",
                    },
                    text: "test",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "padded-numbering-reference",
                    },
                    text: "test",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "padded-numbering-reference",
                    },
                    text: "test",
                }),
            ],
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
