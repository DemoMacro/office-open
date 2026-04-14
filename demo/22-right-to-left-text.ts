// This demo shows right to left for special languages

import * as fs from "fs";

import { Document, Packer, Paragraph, Table, TableCell, TableRow, TextRun } from "docx-plus";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    bidirectional: true,
                    children: [
                        new TextRun({
                            rightToLeft: true,
                            text: "שלום עולם",
                        }),
                    ],
                }),
                new Paragraph({
                    bidirectional: true,
                    children: [
                        new TextRun({
                            bold: true,
                            rightToLeft: true,
                            text: "שלום עולם",
                        }),
                    ],
                }),
                new Paragraph({
                    bidirectional: true,
                    children: [
                        new TextRun({
                            italics: true,
                            rightToLeft: true,
                            text: "שלום עולם",
                        }),
                    ],
                }),
                new Table({
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [new Paragraph("שלום עולם")],
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
                                    children: [new Paragraph("שלום עולם")],
                                }),
                            ],
                        }),
                    ],
                    visuallyRightToLeft: true,
                }),
            ],
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
