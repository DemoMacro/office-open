// Add custom borders to table cell

import * as fs from "fs";

import {
    BorderStyle,
    Document,
    Packer,
    Paragraph,
    Table,
    TableCell,
    TableRow,
} from "@office-open/docx";

const doc = new Document({
    sections: [
        {
            children: [
                new Table({
                    rows: [
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
                                    borders: {
                                        bottom: {
                                            color: "0000FF",
                                            size: 3,
                                            style: BorderStyle.DOUBLE,
                                        },
                                        left: {
                                            color: "00FF00",
                                            size: 3,
                                            style: BorderStyle.DASH_DOT_STROKED,
                                        },
                                        right: {
                                            color: "#ff8000",
                                            size: 3,
                                            style: BorderStyle.DASH_DOT_STROKED,
                                        },
                                        top: {
                                            color: "FF0000",
                                            size: 3,
                                            style: BorderStyle.DASH_DOT_STROKED,
                                        },
                                    },
                                    children: [new Paragraph("Hello")],
                                }),
                                new TableCell({
                                    children: [],
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
                                new TableCell({
                                    children: [],
                                }),
                            ],
                        }),
                    ],
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
