// This demo shows right to left for special languages

import * as fs from "fs";

import {
    Bdo,
    Dir,
    Document,
    Packer,
    Paragraph,
    Table,
    TableCell,
    TableRow,
    TextRun,
} from "docx-plus";

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
                // Bidirectional override: embed LTR text inside RTL paragraph
                new Paragraph({
                    bidirectional: true,
                    children: [
                        new TextRun({
                            rightToLeft: true,
                            text: "مرحبا ",
                        }),
                        new Dir({
                            val: "ltr",
                            children: [new TextRun("Hello World")],
                        }),
                        new TextRun({
                            rightToLeft: true,
                            text: " مرحبا",
                        }),
                    ],
                }),
                // BDO: strong bidirectional override
                new Paragraph({
                    bidirectional: true,
                    children: [
                        new TextRun({
                            rightToLeft: true,
                            text: "نص عربي ",
                        }),
                        new Bdo({
                            val: "ltr",
                            children: [new TextRun("Forced LTR: 123")],
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

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
