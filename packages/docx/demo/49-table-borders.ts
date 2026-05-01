// Add custom borders and no-borders to the table itself

import * as fs from "fs";

import {
    BorderStyle,
    Document,
    HeadingLevel,
    Packer,
    Paragraph,
    Table,
    TableBorders,
    TableCell,
    TableRow,
    TextDirection,
    VerticalAlignTable,
} from "docx-plus";

const table = new Table({
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    borders: {
                        bottom: {
                            color: "ff0000",
                            size: 1,
                            style: BorderStyle.DASH_SMALL_GAP,
                        },
                        left: {
                            color: "ff0000",
                            size: 1,
                            style: BorderStyle.DASH_SMALL_GAP,
                        },
                        right: {
                            color: "ff0000",
                            size: 1,
                            style: BorderStyle.DASH_SMALL_GAP,
                        },
                        top: {
                            color: "ff0000",
                            size: 1,
                            style: BorderStyle.DASH_SMALL_GAP,
                        },
                    },
                    children: [new Paragraph("Hello")],
                }),
                new TableCell({
                    children: [],
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [],
                }),
                new TableCell({
                    children: [new Paragraph("World")],
                }),
            ],
        }),
    ],
});

// Using the no-border convenience object. It is the same as writing this manually:
// Const borders = {
//     Top: {
//         Style: BorderStyle.NONE,
//         Size: 0,
//         Color: "auto",
//     },
//     Bottom: {
//         Style: BorderStyle.NONE,
//         Size: 0,
//         Color: "auto",
//     },
//     Left: {
//         Style: BorderStyle.NONE,
//         Size: 0,
//         Color: "auto",
//     },
//     Right: {
//         Style: BorderStyle.NONE,
//         Size: 0,
//         Color: "auto",
//     },
//     InsideHorizontal: {
//         Style: BorderStyle.NONE,
//         Size: 0,
//         Color: "auto",
//     },
//     InsideVertical: {
//         Style: BorderStyle.NONE,
//         Size: 0,
//         Color: "auto",
//     },
// };
const noBorderTable = new Table({
    borders: TableBorders.NONE,
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
                    children: [new Paragraph({ text: "bottom to top" }), new Paragraph({})],
                    textDirection: TextDirection.BOTTOM_TO_TOP_LEFT_TO_RIGHT,
                }),
                new TableCell({
                    children: [new Paragraph({ text: "top to bottom" }), new Paragraph({})],
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
});

const doc = new Document({
    sections: [{ children: [table, new Paragraph("Hello"), noBorderTable] }],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
