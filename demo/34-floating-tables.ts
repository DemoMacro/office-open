// Example of how you would create a table with float positions

import * as fs from "fs";

import {
    Document,
    OverlapType,
    Packer,
    Paragraph,
    RelativeHorizontalPosition,
    RelativeVerticalPosition,
    Table,
    TableAnchorType,
    TableCell,
    TableLayoutType,
    TableRow,
    WidthType,
} from "docx";

const table = new Table({
    float: {
        bottomFromText: 30,
        horizontalAnchor: TableAnchorType.MARGIN,
        leftFromText: 1000,
        overlap: OverlapType.NEVER,
        relativeHorizontalPosition: RelativeHorizontalPosition.RIGHT,
        relativeVerticalPosition: RelativeVerticalPosition.BOTTOM,
        rightFromText: 2000,
        topFromText: 1500,
        verticalAnchor: TableAnchorType.MARGIN,
    },
    layout: TableLayoutType.FIXED,
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
    width: {
        size: 4535,
        type: WidthType.DXA,
    },
});

const doc = new Document({
    sections: [
        {
            children: [table],
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
