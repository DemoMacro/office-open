// Track Revisions for paragraph properties, section properties, tables
// Docs/usage/change-tracking.md

import * as fs from "fs";

import {
    AlignmentType,
    BorderStyle,
    DeletedTextRun,
    Document,
    HeightRule,
    InsertedTextRun,
    Packer,
    Paragraph,
    Table,
    TableCell,
    TableRow,
    TextRun,
    VerticalAlignTable,
    WidthType,
} from "docx";

const REVISION_DATE = "2020-10-06T09:00:00Z";
const REVISION_AUTHOR = "Firstname Lastname";

const doc = new Document({
    features: {
        trackRevisions: true,
    },
    sections: [
        {
            children: [
                // Paragraph properties revision
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                        new TextRun({
                            text: "Paragraph properties revision",
                            bold: true,
                        }),
                    ],
                    heading: "Heading1",
                    revision: {
                        alignment: AlignmentType.LEFT,
                        author: REVISION_AUTHOR,
                        date: REVISION_DATE,
                        id: 10,
                    },
                }),
                new Paragraph({ text: "" }),

                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "Table 1: Inserted and Deleted Table Row",
                        }),
                    ],
                }),
                new Table({
                    columnWidths: [2000, 2000],
                    layout: "fixed",
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph("Cell 1")] }),
                                new TableCell({ children: [new Paragraph("Cell 2")] }),
                            ],
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [new TextRun("Inserted row cell")],
                                            run: {
                                                insertion: {
                                                    author: REVISION_AUTHOR,
                                                    date: REVISION_DATE,
                                                    id: 0,
                                                },
                                            },
                                        }),
                                    ],
                                }),
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [new TextRun("Inserted row cell")],
                                            run: {
                                                insertion: {
                                                    author: REVISION_AUTHOR,
                                                    date: REVISION_DATE,
                                                    id: 0,
                                                },
                                            },
                                        }),
                                    ],
                                }),
                            ],
                            insertion: {
                                author: REVISION_AUTHOR,
                                date: REVISION_DATE,
                                id: 0,
                            },
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [new TextRun("Deleted row cell")],
                                            run: {
                                                deletion: {
                                                    author: REVISION_AUTHOR,
                                                    date: REVISION_DATE,
                                                    id: 1,
                                                },
                                            },
                                        }),
                                    ],
                                }),
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [new TextRun("Deleted row cell")],
                                            run: {
                                                deletion: {
                                                    author: REVISION_AUTHOR,
                                                    date: REVISION_DATE,
                                                    id: 1,
                                                },
                                            },
                                        }),
                                    ],
                                }),
                            ],
                            deletion: {
                                author: REVISION_AUTHOR,
                                date: REVISION_DATE,
                                id: 1,
                            },
                        }),
                    ],
                }),
                new Paragraph({ text: "" }),

                new Paragraph({
                    children: [
                        new TextRun({ bold: true, text: "Table 2: Inserted  Table Column" }),
                    ],
                }),
                new Table({
                    columnWidths: [2000, 2000],
                    layout: "fixed",
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new InsertedTextRun({
                                                    author: REVISION_AUTHOR,
                                                    date: REVISION_DATE,
                                                    id: 2,
                                                    text: "Inserted cell",
                                                }),
                                            ],
                                            run: {
                                                insertion: {
                                                    author: REVISION_AUTHOR,
                                                    date: REVISION_DATE,
                                                    id: 2,
                                                },
                                            },
                                        }),
                                    ],
                                    insertion: {
                                        author: REVISION_AUTHOR,
                                        date: REVISION_DATE,
                                        id: 2,
                                    },
                                }),
                                new TableCell({ children: [new Paragraph("Cell")] }),
                            ],
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new InsertedTextRun({
                                                    author: REVISION_AUTHOR,
                                                    date: REVISION_DATE,
                                                    id: 2,
                                                    text: "Inserted cell",
                                                }),
                                            ],
                                            run: {
                                                insertion: {
                                                    author: REVISION_AUTHOR,
                                                    date: REVISION_DATE,
                                                    id: 2,
                                                },
                                            },
                                        }),
                                    ],
                                    insertion: {
                                        author: REVISION_AUTHOR,
                                        date: REVISION_DATE,
                                        id: 2,
                                    },
                                }),
                                new TableCell({ children: [new Paragraph("Cell")] }),
                            ],
                        }),
                    ],
                }),
                new Paragraph({ text: "" }),

                new Paragraph({
                    children: [new TextRun({ bold: true, text: "Table 3: Deleted Table Column" })],
                }),
                new Table({
                    columnWidths: [2000, 2000],
                    layout: "fixed",
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph("Cell")] }),
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new DeletedTextRun({
                                                    author: REVISION_AUTHOR,
                                                    date: REVISION_DATE,
                                                    id: 3,
                                                    text: "Deleted cell",
                                                }),
                                            ],
                                            run: {
                                                deletion: {
                                                    author: REVISION_AUTHOR,
                                                    date: REVISION_DATE,
                                                    id: 3,
                                                },
                                            },
                                        }),
                                    ],
                                    deletion: {
                                        author: REVISION_AUTHOR,
                                        date: REVISION_DATE,
                                        id: 3,
                                    },
                                }),
                            ],
                        }),
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph("Cell")] }),
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new DeletedTextRun({
                                                    author: REVISION_AUTHOR,
                                                    date: REVISION_DATE,
                                                    id: 3,
                                                    text: "Deleted cell",
                                                }),
                                            ],
                                            run: {
                                                deletion: {
                                                    author: REVISION_AUTHOR,
                                                    date: REVISION_DATE,
                                                    id: 3,
                                                },
                                            },
                                        }),
                                    ],
                                    deletion: {
                                        author: REVISION_AUTHOR,
                                        date: REVISION_DATE,
                                        id: 3,
                                    },
                                }),
                            ],
                        }),
                    ],
                }),
                new Paragraph({ text: "" }),

                new Paragraph({
                    children: [
                        new TextRun({ bold: true, text: "Table 3: cell properties revision" }),
                    ],
                }),
                new Table({
                    columnWidths: [2000, 2000],
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    width: { size: 4000, type: WidthType.DXA },
                                    shading: {
                                        color: "#00FF00",
                                        fill: "#00FF00",
                                    },
                                    verticalAlign: VerticalAlignTable.CENTER,
                                    revision: {
                                        width: { size: 2000, type: WidthType.DXA },
                                        id: 4,
                                        author: REVISION_AUTHOR,
                                        date: REVISION_DATE,
                                        verticalAlign: VerticalAlignTable.TOP,
                                    },
                                    children: [new Paragraph("Cell 1")],
                                }),
                                new TableCell({
                                    width: { size: 2000, type: WidthType.DXA },
                                    shading: {
                                        color: "#00FF00",
                                        fill: "#00FF00",
                                    },
                                    revision: {
                                        width: { size: 2000, type: WidthType.DXA },
                                        id: 4,
                                        author: REVISION_AUTHOR,
                                        date: REVISION_DATE,
                                    },
                                    children: [new Paragraph("Cell 2")],
                                }),
                            ],
                            height: { rule: HeightRule.EXACT, value: 600 },
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [new Paragraph("Cell 3")],
                                    revision: {
                                        author: REVISION_AUTHOR,
                                        date: REVISION_DATE,
                                        id: 4,
                                        width: { size: 2000, type: WidthType.DXA },
                                    },
                                    width: { size: 4000, type: WidthType.DXA },
                                }),
                                new TableCell({
                                    children: [new Paragraph("Cell 4")],
                                    revision: {
                                        author: REVISION_AUTHOR,
                                        date: REVISION_DATE,
                                        id: 4,
                                        width: { size: 2000, type: WidthType.DXA },
                                    },
                                    width: { size: 2000, type: WidthType.DXA },
                                }),
                            ],
                        }),
                    ],
                }),
                new Paragraph({ text: "" }),

                new Paragraph({
                    children: [
                        new TextRun({ bold: true, text: "Table 4: row properties revision" }),
                    ],
                }),
                new Table({
                    columnWidths: [2000, 2000],
                    layout: "fixed",
                    rows: [
                        new TableRow({
                            cantSplit: true,
                            children: [
                                new TableCell({ children: [new Paragraph("Cell 1")] }),
                                new TableCell({ children: [new Paragraph("Cell 2")] }),
                            ],
                            height: { rule: HeightRule.EXACT, value: 600 },
                            revision: {
                                author: REVISION_AUTHOR,
                                cantSplit: false,
                                date: REVISION_DATE,
                                id: 5,
                                tableHeader: false,
                            },
                            tableHeader: true,
                        }),
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph("Cell 3")] }),
                                new TableCell({ children: [new Paragraph("Cell 4")] }),
                            ],
                        }),
                    ],
                }),
                new Paragraph({ text: "" }),

                new Table({
                    borders: {
                        bottom: {
                            color: "FF0000",
                            size: 5,
                            style: BorderStyle.DASHED,
                        },
                        left: {
                            color: "FF0000",
                            size: 5,
                            style: BorderStyle.DASHED,
                        },
                        right: {
                            color: "FF0000",
                            size: 5,
                            style: BorderStyle.DASHED,
                        },
                        top: {
                            color: "FF0000",
                            size: 5,
                            style: BorderStyle.DASHED,
                        },
                    },
                    columnWidths: [2000, 2000],
                    layout: "fixed",
                    revision: {
                        author: REVISION_AUTHOR,
                        borders: {
                            bottom: {
                                color: "00FF00",
                                size: 3,
                                style: BorderStyle.DOTTED,
                            },
                            left: {
                                color: "00FF00",
                                size: 3,
                                style: BorderStyle.DOTTED,
                            },
                            right: {
                                color: "00FF00",
                                size: 3,
                                style: BorderStyle.DOTTED,
                            },
                            top: {
                                color: "00FF00",
                                size: 3,
                                style: BorderStyle.DOTTED,
                            },
                        },
                        date: REVISION_DATE,
                        id: 6,
                    },
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph("Cell 1")] }),
                                new TableCell({ children: [new Paragraph("Cell 2")] }),
                            ],
                        }),
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph("Cell 3")] }),
                                new TableCell({ children: [new Paragraph("Cell 4")] }),
                            ],
                        }),
                    ],
                }),
                new Paragraph({ text: "" }),
            ],
            properties: {},
        },
        // Section properties revision
        {
            children: [
                new Paragraph({ text: "Paragraph in different section with revision properties" }),
            ],
            properties: {
                page: {
                    margin: {
                        bottom: 2440,
                        left: 2440,
                        right: 2440,
                        top: 2440,
                    },
                    size: {
                        height: 11_909,
                        width: 16_834,
                    },
                },
                revision: {
                    author: REVISION_AUTHOR,
                    date: REVISION_DATE,
                    id: 20,
                    page: {
                        margin: {
                            bottom: 1440,
                            left: 1440,
                            right: 1440,
                            top: 1440,
                        },
                    },
                },
            },
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("98-track-revisions-2.docx", buffer);
    console.log("Document created successfully at 98-track-revisions-2.docx");
});
