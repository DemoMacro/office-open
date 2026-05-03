// Setting styles with JavaScript configuration

import * as fs from "fs";

import {
    AlignmentType,
    Document,
    Footer,
    HeadingLevel,
    ImageRun,
    LevelFormat,
    Packer,
    Paragraph,
    TabStopPosition,
    Table,
    TableCell,
    TableRow,
    UnderlineType,
    convertInchesToTwip,
} from "@office-open/docx";

const table = new Table({
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("Test cell 1.")],
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("Test cell 2.")],
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("Test cell 3.")],
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("Test cell 4.")],
                }),
            ],
        }),
    ],
});

const doc = new Document({
    numbering: {
        config: [
            {
                levels: [
                    {
                        level: 0,
                        format: LevelFormat.DECIMAL,
                        text: "%1)",
                        start: 50,
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
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("./demo/images/pizza.gif"),
                            transformation: {
                                width: 100,
                                height: 100,
                            },
                            type: "gif",
                        }),
                    ],
                }),
                new Paragraph({
                    text: "HEADING",
                    heading: HeadingLevel.HEADING_1,
                    alignment: AlignmentType.CENTER,
                }),
                new Paragraph({
                    text: "Ref. :",
                    style: "normalPara",
                }),
                new Paragraph({
                    text: "Date :",
                    style: "normalPara",
                }),
                new Paragraph({
                    text: "To,",
                    style: "normalPara",
                }),
                new Paragraph({
                    text: "The Superindenting Engineer,(O &M)",
                    style: "normalPara",
                }),
                new Paragraph({
                    text: "Sub : ",
                    style: "normalPara",
                }),
                new Paragraph({
                    text: "Ref. : ",
                    style: "normalPara",
                }),
                new Paragraph({
                    text: "Sir,",
                    style: "normalPara",
                }),
                new Paragraph({
                    text: "BRIEF DESCRIPTION",
                    style: "normalPara",
                }),
                table,
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("./demo/images/pizza.gif"),
                            transformation: {
                                width: 100,
                                height: 100,
                            },
                            type: "gif",
                        }),
                    ],
                }),
                new Paragraph({
                    text: "Test",
                    style: "normalPara2",
                }),
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("./demo/images/pizza.gif"),
                            transformation: {
                                width: 100,
                                height: 100,
                            },
                            type: "gif",
                        }),
                    ],
                }),
                new Paragraph({
                    text: "Test 2",
                    style: "normalPara2",
                }),
                new Paragraph({
                    text: "Numbered paragraph that has numbering attached to custom styles",
                    style: "numberedPara",
                }),
                new Paragraph({
                    text: "Numbered para would show up in the styles pane at Word",
                    style: "numberedPara",
                }),
            ],
            footers: {
                default: new Footer({
                    children: [
                        new Paragraph({
                            text: "1",
                            style: "normalPara",
                            alignment: AlignmentType.RIGHT,
                        }),
                    ],
                }),
            },
            properties: {
                page: {
                    margin: {
                        bottom: 700,
                        left: 700,
                        right: 700,
                        top: 700,
                    },
                },
            },
        },
    ],
    styles: {
        default: {
            heading1: {
                paragraph: {
                    alignment: AlignmentType.CENTER,
                    spacing: { line: 340 },
                },
                run: {
                    bold: true,
                    color: "000000",
                    font: "Calibri",
                    size: 52,
                    underline: {
                        color: "000000",
                        type: UnderlineType.SINGLE,
                    },
                },
            },
            heading2: {
                paragraph: {
                    spacing: { line: 340 },
                },
                run: {
                    bold: true,
                    font: "Calibri",
                    size: 26,
                },
            },
            heading3: {
                paragraph: {
                    spacing: { line: 276 },
                },
                run: {
                    bold: true,
                    font: "Calibri",
                    size: 26,
                },
            },
            heading4: {
                paragraph: {
                    alignment: AlignmentType.JUSTIFIED,
                },
                run: {
                    bold: true,
                    font: "Calibri",
                    size: 26,
                },
            },
        },
        paragraphStyles: [
            {
                basedOn: "Normal",
                id: "normalPara",
                name: "Normal Para",
                next: "Normal",
                paragraph: {
                    leftTabStop: 453.543_307_087,
                    rightTabStop: TabStopPosition.MAX,
                    spacing: { after: 20 * 72 * 0.05, before: 20 * 72 * 0.1, line: 276 },
                },
                quickFormat: true,
                run: {
                    bold: true,
                    font: "Calibri",
                    size: 26,
                },
            },
            {
                basedOn: "Normal",
                id: "normalPara2",
                name: "Normal Para2",
                next: "Normal",
                paragraph: {
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { after: 20 * 72 * 0.05, before: 20 * 72 * 0.1, line: 276 },
                },
                quickFormat: true,
                run: {
                    font: "Calibri",
                    size: 26,
                },
            },
            {
                basedOn: "Normal",
                id: "aside",
                name: "Aside",
                next: "Normal",
                paragraph: {
                    indent: { left: convertInchesToTwip(0.5) },
                    spacing: { line: 276 },
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
            },
            {
                basedOn: "Normal",
                id: "numberedPara",
                name: "Numbered Para",
                next: "Normal",
                paragraph: {
                    leftTabStop: 453.543_307_087,
                    numbering: {
                        instance: 0,
                        level: 0,
                        reference: "ref1",
                    },
                    rightTabStop: TabStopPosition.MAX,
                    spacing: { after: 20 * 72 * 0.05, before: 20 * 72 * 0.1, line: 276 },
                },
                quickFormat: true,
                run: {
                    bold: true,
                    font: "Calibri",
                    size: 26,
                },
            },
        ],
    },
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
