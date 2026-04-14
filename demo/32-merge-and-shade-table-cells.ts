// Example of how you would merge cells together (Rows and Columns) and apply shading
// Also includes an example on how to center tables

import * as fs from "fs";

import {
    AlignmentType,
    BorderStyle,
    Document,
    HeadingLevel,
    Packer,
    Paragraph,
    ShadingType,
    Table,
    TableCell,
    TableRow,
    WidthType,
    convertInchesToTwip,
} from "docx";

const table = new Table({
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("Hello")],
                    columnSpan: 2,
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [],
                }),
                new TableCell({
                    children: [],
                }),
            ],
        }),
    ],
});

const table2 = new Table({
    alignment: AlignmentType.CENTER,
    columnWidths: [convertInchesToTwip(0.69), convertInchesToTwip(0.69), convertInchesToTwip(0.69)],
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("World")],
                    columnSpan: 3,
                    margins: {
                        bottom: convertInchesToTwip(0.69),
                        left: convertInchesToTwip(0.69),
                        right: convertInchesToTwip(0.69),
                        top: convertInchesToTwip(0.69),
                    },
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [],
                }),
                new TableCell({
                    children: [],
                }),
                new TableCell({
                    children: [],
                }),
            ],
        }),
    ],
    width: {
        size: 100,
        type: WidthType.AUTO,
    },
});

const table3 = new Table({
    alignment: AlignmentType.CENTER,
    margins: {
        bottom: convertInchesToTwip(0.27),
        left: convertInchesToTwip(0.27),
        right: convertInchesToTwip(0.27),
        top: convertInchesToTwip(0.27),
    },
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("Foo")],
                }),
                new TableCell({
                    children: [new Paragraph("v")],
                    columnSpan: 3,
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("Bar1")],
                    shading: {
                        color: "auto",
                        fill: "b79c2f",
                        type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                    },
                }),
                new TableCell({
                    children: [new Paragraph("Bar2")],
                    shading: {
                        color: "auto",
                        fill: "42c5f4",
                        type: ShadingType.PERCENT_95,
                    },
                }),
                new TableCell({
                    children: [new Paragraph("Bar3")],
                    shading: {
                        color: "e2df0b",
                        fill: "880aa8",
                        type: ShadingType.PERCENT_10,
                    },
                }),
                new TableCell({
                    children: [new Paragraph("Bar4")],
                    shading: {
                        color: "auto",
                        fill: "FF0000",
                        type: ShadingType.CLEAR,
                    },
                }),
            ],
        }),
    ],
    width: {
        size: convertInchesToTwip(4.86),
        type: WidthType.DXA,
    },
});

const table4 = new Table({
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("0,0")],
                    columnSpan: 2,
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("1,0")],
                }),
                new TableCell({
                    children: [new Paragraph("1,1")],
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("2,0")],
                    columnSpan: 2,
                }),
            ],
        }),
    ],
    width: {
        size: 100,
        type: WidthType.PERCENTAGE,
    },
});

const table5 = new Table({
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("0,0")],
                }),
                new TableCell({
                    children: [new Paragraph("0,1")],
                    rowSpan: 2,
                }),
                new TableCell({
                    children: [new Paragraph("0,2")],
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("1,0")],
                }),
                new TableCell({
                    children: [new Paragraph("1,2")],
                    rowSpan: 2,
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("2,0")],
                }),
                new TableCell({
                    children: [new Paragraph("2,1")],
                }),
            ],
        }),
    ],
    width: {
        size: 100,
        type: WidthType.PERCENTAGE,
    },
});

const borders = {
    bottom: {
        color: "FF0000",
        size: 1,
        style: BorderStyle.DASH_SMALL_GAP,
    },
    left: {
        color: "FF0000",
        size: 1,
        style: BorderStyle.DASH_SMALL_GAP,
    },
    right: {
        color: "FF0000",
        size: 1,
        style: BorderStyle.DASH_SMALL_GAP,
    },
    top: {
        color: "FF0000",
        size: 1,
        style: BorderStyle.DASH_SMALL_GAP,
    },
};

const table6 = new Table({
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    borders,
                    children: [new Paragraph("0,0")],
                    rowSpan: 2,
                }),
                new TableCell({
                    borders,
                    children: [new Paragraph("0,1")],
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    borders,
                    children: [new Paragraph("1,1")],
                    rowSpan: 2,
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    borders,
                    children: [new Paragraph("2,0")],
                }),
            ],
        }),
    ],
    width: {
        size: 100,
        type: WidthType.PERCENTAGE,
    },
});

const table7 = new Table({
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("0,0")],
                }),
                new TableCell({
                    children: [new Paragraph("0,1")],
                }),
                new TableCell({
                    children: [new Paragraph("0,2")],
                    rowSpan: 2,
                }),
                new TableCell({
                    children: [new Paragraph("0,3")],
                    rowSpan: 3,
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("1,0")],
                    columnSpan: 2,
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("2,0")],
                    columnSpan: 2,
                }),
                new TableCell({
                    children: [new Paragraph("2,2")],
                    rowSpan: 2,
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("3,0")],
                }),
                new TableCell({
                    children: [new Paragraph("3,1")],
                }),
                new TableCell({
                    children: [new Paragraph("3,3")],
                }),
            ],
        }),
    ],
    width: {
        size: 100,
        type: WidthType.PERCENTAGE,
    },
});

const table8 = new Table({
    rows: [
        new TableRow({
            children: [
                new TableCell({ children: [new Paragraph("1,1")] }),
                new TableCell({ children: [new Paragraph("1,2")] }),
                new TableCell({ children: [new Paragraph("1,3")] }),
                new TableCell({ borders, children: [new Paragraph("1,4")], rowSpan: 4 }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({ children: [new Paragraph("2,1")] }),
                new TableCell({ children: [new Paragraph("2,2")] }),
                new TableCell({ children: [new Paragraph("2,3")], rowSpan: 3 }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({ children: [new Paragraph("3,1")] }),
                new TableCell({ children: [new Paragraph("3,2")], rowSpan: 2 }),
            ],
        }),
        new TableRow({
            children: [new TableCell({ children: [new Paragraph("4,1")] })],
        }),
    ],
    width: {
        size: 100,
        type: WidthType.PERCENTAGE,
    },
});
const doc = new Document({
    sections: [
        {
            children: [
                table,
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    text: "Another table",
                }),
                table2,
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    text: "Another table",
                }),
                table3,
                new Paragraph("Merging columns 1"),
                table4,
                new Paragraph("Merging columns 2"),
                table5,
                new Paragraph("Merging columns 3"),
                table6,
                new Paragraph("Merging columns 4"),
                table7,
                new Paragraph("Merging columns 5"),
                table8,
            ],
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
