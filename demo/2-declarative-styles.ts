// Example on how to customize the look at feel using Styles

import * as fs from "fs";

import {
    AlignmentType,
    Document,
    HeadingLevel,
    LevelFormat,
    Packer,
    Paragraph,
    TextRun,
    UnderlineType,
    convertInchesToTwip,
} from "docx-plus";

const doc = new Document({
    creator: "Clippy",
    description: "A brief example of using docx",
    numbering: {
        config: [
            {
                levels: [
                    {
                        level: 0,
                        format: LevelFormat.LOWER_LETTER,
                        text: "%1)",
                        alignment: AlignmentType.LEFT,
                    },
                ],
                reference: "my-crazy-numbering",
            },
        ],
    },
    sections: [
        {
            children: [
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    text: "Test heading1, bold and italicized",
                }),
                new Paragraph("Some simple content"),
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    text: "Test heading2 with double red underline",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "my-crazy-numbering",
                    },
                    style: "aside",
                    text: "Option1",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "my-crazy-numbering",
                    },
                    text: "Option5 -- override 2 to 5",
                }),
                new Paragraph({
                    numbering: {
                        level: 0,
                        reference: "my-crazy-numbering",
                    },
                    text: "Option3",
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            font: {
                                name: "Monospace",
                            },
                            text: "Some monospaced content",
                        }),
                    ],
                }),
                new Paragraph({
                    style: "aside",
                    text: "An aside, in light gray italics and indented",
                }),
                new Paragraph({
                    style: "wellSpaced",
                    text: "This is normal, but well-spaced text",
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "This is a bold run,",
                        }),
                        new TextRun(" switching to normal "),
                        new TextRun({
                            text: "and then underlined ",
                            underline: {},
                        }),
                        new TextRun({
                            emphasisMark: {},
                            text: "and then emphasis-mark ",
                        }),
                        new TextRun({
                            text: "and back to normal.",
                        }),
                        new TextRun({
                            text: "This text will be invisible!",
                            vanish: true,
                        }),
                        new TextRun({
                            specVanish: true,
                            text: "This text will be VERY invisible! Word processors cannot override this!",
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "Strong Style",
                        }),
                        new TextRun({
                            text: " - Very strong.",
                        }),
                    ],
                    style: "Strong",
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "Underline and Strike",
                        }),
                        new TextRun({
                            text: " Override Underline ",
                            underline: {
                                type: UnderlineType.NONE,
                            },
                        }),
                        new TextRun({
                            text: "Strike and Underline",
                        }),
                    ],
                    style: "strikeUnderline",
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "Hello World ",
                        }),
                        new TextRun({
                            style: "strikeUnderlineCharacter",
                            text: "Underline and Strike",
                        }),
                        new TextRun({
                            text: " Another Hello World",
                        }),
                        new TextRun({
                            scale: 50,
                            text: " Scaled text",
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "Scaled paragraph",
                        }),
                    ],
                }),
            ],
        },
    ],
    styles: {
        characterStyles: [
            {
                basedOn: "Normal",
                id: "strikeUnderlineCharacter",
                name: "Strike Underline",
                quickFormat: true,
                run: {
                    strike: true,
                    underline: {
                        type: UnderlineType.SINGLE,
                    },
                },
            },
        ],
        default: {
            document: {
                paragraph: {
                    alignment: AlignmentType.RIGHT,
                },
                run: {
                    font: "Calibri",
                    size: "11pt",
                },
            },
            heading1: {
                paragraph: {
                    spacing: {
                        after: 120,
                    },
                },
                run: {
                    bold: true,
                    color: "FF0000",
                    italics: true,
                    size: 28,
                },
            },
            heading2: {
                paragraph: {
                    spacing: {
                        after: 120,
                        before: 240,
                    },
                },
                run: {
                    bold: true,
                    size: 26,
                    underline: {
                        color: "FF0000",
                        type: UnderlineType.DOUBLE,
                    },
                },
            },
            listParagraph: {
                run: {
                    color: "#FF0000",
                },
            },
        },
        paragraphStyles: [
            {
                basedOn: "Normal",
                id: "aside",
                name: "Aside",
                next: "Normal",
                paragraph: {
                    indent: {
                        left: convertInchesToTwip(0.5),
                    },
                    spacing: {
                        line: 276,
                    },
                },
                run: {
                    color: "999999",
                    italics: true,
                },
            },
            {
                basedOn: "Normal",
                id: "wellSpaced",
                name: "Well Spaced",
                paragraph: {
                    spacing: { after: 20 * 72 * 0.05, before: 20 * 72 * 0.1, line: 276 },
                },
                quickFormat: true,
            },
            {
                basedOn: "Normal",
                id: "strikeUnderline",
                name: "Strike Underline",
                quickFormat: true,
                run: {
                    strike: true,
                    underline: {
                        type: UnderlineType.SINGLE,
                    },
                },
            },
        ],
    },
    title: "Sample Document",
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
