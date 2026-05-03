// Numbering and bullet points example

import * as fs from "fs";

import {
    AlignmentType,
    Document,
    Footer,
    Header,
    HeadingLevel,
    LevelFormat,
    Packer,
    Paragraph,
    convertInchesToTwip,
} from "@office-open/docx";

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
                    {
                        alignment: AlignmentType.START,
                        format: LevelFormat.DECIMAL,
                        level: 1,
                        style: {
                            paragraph: {
                                indent: {
                                    hanging: convertInchesToTwip(0.68),
                                    left: convertInchesToTwip(1),
                                },
                            },
                        },
                        text: "%2.",
                    },
                    {
                        alignment: AlignmentType.START,
                        format: LevelFormat.LOWER_LETTER,
                        level: 2,
                        style: {
                            paragraph: {
                                indent: {
                                    hanging: convertInchesToTwip(1.18),
                                    left: convertInchesToTwip(1.5),
                                },
                            },
                        },
                        text: "%3)",
                    },
                    {
                        alignment: AlignmentType.START,
                        format: LevelFormat.UPPER_LETTER,
                        level: 3,
                        style: {
                            paragraph: {
                                indent: { hanging: 2420, left: 2880 },
                            },
                        },
                        text: "%4)",
                    },
                ],
                reference: "my-crazy-numbering",
            },
            {
                levels: [
                    {
                        alignment: AlignmentType.LEFT,
                        format: LevelFormat.BULLET,
                        level: 0,
                        style: {
                            paragraph: {
                                indent: {
                                    hanging: convertInchesToTwip(0.25),
                                    left: convertInchesToTwip(0.5),
                                },
                            },
                        },
                        text: "\u1F60",
                    },
                    {
                        alignment: AlignmentType.LEFT,
                        format: LevelFormat.BULLET,
                        level: 1,
                        style: {
                            paragraph: {
                                indent: {
                                    hanging: convertInchesToTwip(0.25),
                                    left: convertInchesToTwip(1),
                                },
                            },
                        },
                        text: "\u00A5",
                    },
                    {
                        alignment: AlignmentType.LEFT,
                        format: LevelFormat.BULLET,
                        level: 2,
                        style: {
                            paragraph: {
                                indent: { hanging: convertInchesToTwip(0.25), left: 2160 },
                            },
                        },
                        text: "\u273F",
                    },
                    {
                        alignment: AlignmentType.LEFT,
                        format: LevelFormat.BULLET,
                        level: 3,
                        style: {
                            paragraph: {
                                indent: { hanging: convertInchesToTwip(0.25), left: 2880 },
                            },
                        },
                        text: "\u267A",
                    },
                    {
                        alignment: AlignmentType.LEFT,
                        format: LevelFormat.BULLET,
                        level: 4,
                        style: {
                            paragraph: {
                                indent: { hanging: convertInchesToTwip(0.25), left: 3600 },
                            },
                        },
                        text: "\u2603",
                    },
                ],
                reference: "my-unique-bullet-points",
            },
        ],
    },
    sections: [
        {
            children: [
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "my-crazy-numbering",
                    },
                    text: "Hey you",
                }),
                new Paragraph({
                    numbering: {
                        level: 1,
                        reference: "my-crazy-numbering",
                    },
                    text: "What's up fam",
                }),
                new Paragraph({
                    numbering: {
                        level: 1,
                        reference: "my-crazy-numbering",
                    },
                    text: "Hello World 2",
                }),
                new Paragraph({
                    numbering: {
                        level: 2,
                        reference: "my-crazy-numbering",
                    },
                    text: "Yeah boi",
                }),
                new Paragraph({
                    bullet: {
                        level: 0,
                    },
                    text: "Hey you",
                }),
                new Paragraph({
                    bullet: {
                        level: 1,
                    },
                    text: "What's up fam",
                }),
                new Paragraph({
                    bullet: {
                        level: 2,
                    },
                    text: "Hello World 2",
                }),
                new Paragraph({
                    bullet: {
                        level: 3,
                    },
                    text: "Yeah boi",
                }),
                new Paragraph({
                    numbering: {
                        level: 3,
                        reference: "my-crazy-numbering",
                    },
                    text: "101 MSXFM",
                }),
                new Paragraph({
                    numbering: {
                        level: 1,
                        reference: "my-crazy-numbering",
                    },
                    text: "back to level 1",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "my-crazy-numbering",
                    },
                    text: "back to level 0",
                }),

                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    text: "Custom Bullet points",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "my-unique-bullet-points",
                    },
                    text: "What's up fam",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "my-unique-bullet-points",
                    },
                    text: "Hey you",
                }),
                new Paragraph({
                    numbering: {
                        level: 1,
                        reference: "my-unique-bullet-points",
                    },
                    text: "What's up fam",
                }),
                new Paragraph({
                    numbering: {
                        level: 2,
                        reference: "my-unique-bullet-points",
                    },
                    text: "Hello World 2",
                }),
                new Paragraph({
                    numbering: {
                        level: 3,
                        reference: "my-unique-bullet-points",
                    },
                    text: "Yeah boi",
                }),
                new Paragraph({
                    numbering: {
                        level: 4,
                        reference: "my-unique-bullet-points",
                    },
                    text: "my Awesome numbering",
                }),
                new Paragraph({
                    numbering: {
                        level: 1,
                        reference: "my-unique-bullet-points",
                    },
                    text: "Back to level 1",
                }),
            ],
            footers: {
                default: new Footer({
                    children: [
                        new Paragraph({
                            numbering: {
                                level: 0,
                                reference: "my-crazy-numbering",
                            },
                            text: "Hey you",
                        }),
                        new Paragraph({
                            numbering: {
                                level: 1,
                                reference: "my-crazy-numbering",
                            },
                            text: "What's up fam",
                        }),
                    ],
                }),
            },
            headers: {
                default: new Header({
                    children: [
                        new Paragraph({
                            numbering: {
                                level: 0,
                                reference: "my-crazy-numbering",
                            },
                            text: "Hey you",
                        }),
                        new Paragraph({
                            numbering: {
                                level: 1,
                                reference: "my-crazy-numbering",
                            },
                            text: "What's up fam",
                        }),
                    ],
                }),
            },
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
