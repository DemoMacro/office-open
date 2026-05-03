// Example of how you would create a table and add data to it

import * as fs from "fs";

import {
    Document,
    Packer,
    Paragraph,
    Table,
    TableCell,
    TableRow,
    WidthType,
} from "@office-open/docx";

const table = new Table({
    columnWidths: [3505, 5505],
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("Hello")],
                    width: {
                        size: 3505,
                        type: WidthType.DXA,
                    },
                }),
                new TableCell({
                    children: [],
                    width: {
                        size: 5505,
                        type: WidthType.DXA,
                    },
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [],
                    width: {
                        size: 3505,
                        type: WidthType.DXA,
                    },
                }),
                new TableCell({
                    children: [new Paragraph("World")],
                    width: {
                        size: 5505,
                        type: WidthType.DXA,
                    },
                }),
            ],
        }),
    ],
});

const table2 = new Table({
    columnWidths: [4505, 4505],
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("Hello")],
                    width: {
                        size: 4505,
                        type: WidthType.DXA,
                    },
                }),
                new TableCell({
                    children: [],
                    width: {
                        size: 4505,
                        type: WidthType.DXA,
                    },
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [],
                    width: {
                        size: 4505,
                        type: WidthType.DXA,
                    },
                }),
                new TableCell({
                    children: [new Paragraph("World")],
                    width: {
                        size: 4505,
                        type: WidthType.DXA,
                    },
                }),
            ],
        }),
    ],
});

const table3 = new Table({
    rows: [
        new TableRow({
            children: [
                new TableCell({
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

// Table with new properties: row cnfStyle, gridBefore/gridAfter, rowAlignment,
// cell noWrap, fitText, horizontalMerge, headers
const table4 = new Table({
    styleRowBandSize: 3,
    styleColBandSize: 2,
    caption: "Sales Report",
    description: "Quarterly sales data",
    rows: [
        new TableRow({
            cnfStyle: { val: "000000000010" },
            gridBefore: 1,
            gridAfter: 0,
            rowAlignment: "center" as const,
            children: [
                new TableCell({
                    children: [new Paragraph("Header A")],
                    noWrap: true,
                    fitText: true,
                }),
                new TableCell({
                    children: [new Paragraph("Header B")],
                    noWrap: true,
                    fitText: true,
                }),
            ],
        }),
        new TableRow({
            widthBefore: { size: 100, type: WidthType.DXA },
            widthAfter: { size: 100, type: WidthType.DXA },
            children: [
                new TableCell({
                    children: [new Paragraph("Cell with long text that should not wrap")],
                    noWrap: true,
                }),
                new TableCell({
                    children: [new Paragraph("Normal cell")],
                    horizontalMerge: "continue",
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("Restart merge")],
                    horizontalMerge: "restart",
                }),
                new TableCell({
                    children: [new Paragraph("Continued")],
                    horizontalMerge: "continue",
                }),
            ],
        }),
    ],
});

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({ text: "Table with skewed widths" }),
                table,
                new Paragraph({ text: "Table with equal widths" }),
                table2,
                new Paragraph({ text: "Table without setting widths" }),
                table3,
                new Paragraph({
                    text: "Table with new properties (cnfStyle, gridBefore/After, noWrap, fitText, horizontalMerge, caption, description)",
                }),
                table4,
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
