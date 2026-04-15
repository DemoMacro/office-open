// Patch a document with patches

import * as fs from "fs";

import {
    ExternalHyperlink,
    HeadingLevel,
    ImageRun,
    Paragraph,
    PatchType,
    Table,
    TableCell,
    TableRow,
    TextDirection,
    TextRun,
    VerticalAlignTable,
    patchDocument,
} from "docx-plus";

const doc = await patchDocument({
    data: fs.readFileSync("demo/assets/simple-template-4.docx"),
    outputType: "nodebuffer",
    patches: {
        footer_text: {
            children: [
                new TextRun("replaced just as"),
                new TextRun(" well"),
                new ExternalHyperlink({
                    children: [
                        new TextRun({
                            text: "BBC News Link",
                        }),
                    ],
                    link: "https://www.bbc.co.uk/news",
                }),
            ],
            type: PatchType.PARAGRAPH,
        },
        header_adjective: {
            children: [new TextRun("Delightful Header")],
            type: PatchType.PARAGRAPH,
        },
        image_test: {
            children: [
                new ImageRun({
                    data: fs.readFileSync("./demo/images/image1.jpeg"),
                    transformation: { height: 100, width: 100 },
                    type: "jpg",
                }),
            ],
            type: PatchType.PARAGRAPH,
        },
        item_1: {
            children: [
                new TextRun("#657"),
                new ExternalHyperlink({
                    children: [
                        new TextRun({
                            text: "BBC News Link",
                        }),
                    ],
                    link: "https://www.bbc.co.uk/news",
                }),
            ],
            type: PatchType.PARAGRAPH,
        },
        name: {
            children: [
                new TextRun("Sir. "),
                new TextRun("John Doe"),
                new TextRun("(The Conqueror)"),
            ],
            type: PatchType.PARAGRAPH,
        },
        paragraph_replace: {
            children: [
                new Paragraph("Lorem ipsum paragraph"),
                new Paragraph("Another paragraph"),
                new Paragraph({
                    children: [
                        new TextRun("This is a "),
                        new ExternalHyperlink({
                            children: [
                                new TextRun({
                                    text: "Google Link",
                                }),
                            ],
                            link: "https://www.google.co.uk",
                        }),
                        new ImageRun({
                            data: fs.readFileSync("./demo/images/dog.png"),
                            transformation: { height: 100, width: 100 },
                            type: "png",
                        }),
                    ],
                }),
            ],
            type: PatchType.DOCUMENT,
        },
        table: {
            children: [
                new Table({
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [new Paragraph({}), new Paragraph({})],
                                    verticalAlign: VerticalAlignTable.CENTER,
                                }),
                                new TableCell({
                                    children: [new Paragraph({}), new Paragraph({})],
                                    verticalAlign: VerticalAlignTable.CENTER,
                                }),
                                new TableCell({
                                    children: [
                                        new Paragraph({ text: "bottom to top" }),
                                        new Paragraph({}),
                                    ],
                                    textDirection: TextDirection.BOTTOM_TO_TOP_LEFT_TO_RIGHT,
                                }),
                                new TableCell({
                                    children: [
                                        new Paragraph({ text: "top to bottom" }),
                                        new Paragraph({}),
                                    ],
                                    textDirection: TextDirection.TOP_TO_BOTTOM_RIGHT_TO_LEFT,
                                }),
                            ],
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            heading: HeadingLevel.HEADING_1,
                                            text: "Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah",
                                        }),
                                    ],
                                }),
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            text: "This text should be in the middle of the cell",
                                        }),
                                    ],
                                    verticalAlign: VerticalAlignTable.CENTER,
                                }),
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            text: "Text above should be vertical from bottom to top",
                                        }),
                                    ],
                                    verticalAlign: VerticalAlignTable.CENTER,
                                }),
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            text: "Text above should be vertical from top to bottom",
                                        }),
                                    ],
                                    verticalAlign: VerticalAlignTable.CENTER,
                                }),
                            ],
                        }),
                    ],
                }),
            ],
            type: PatchType.DOCUMENT,
        },
        table_heading_1: {
            children: [new TextRun("Heading wow!")],
            type: PatchType.PARAGRAPH,
        },
    },
    placeholderDelimiters: { end: ">>", start: "<<" },
});
fs.writeFileSync("My Document.docx", doc);
