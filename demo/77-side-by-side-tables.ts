// Exporting the document as a stream

import * as fs from "fs";

import {
    Document,
    Packer,
    Paragraph,
    Table,
    TableBorders,
    TableCell,
    TableRow,
    WidthType,
} from "docx";

const table1 = new Table({
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
    columnWidths: [3505, 5505],
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("Foo")],
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
                    children: [new Paragraph("Bar")],
                    width: {
                        size: 5505,
                        type: WidthType.DXA,
                    },
                }),
            ],
        }),
    ],
});

const noBorderTable = new Table({
    borders: TableBorders.NONE,
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    children: [table1],
                }),
                new TableCell({
                    children: [table2],
                }),
            ],
        }),
    ],
});

const doc = new Document({
    sections: [
        {
            children: [noBorderTable],
            properties: {},
        },
    ],
});

const stream = Packer.toStream(doc);
stream.pipe(fs.createWriteStream("My Document.docx"));
