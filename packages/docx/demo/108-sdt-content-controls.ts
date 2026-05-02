// Demo: Structured Document Tags (SDT) - content controls
import * as fs from "fs";

import {
    Document,
    Packer,
    Paragraph,
    StructuredDocumentTagBlock,
    StructuredDocumentTagCell,
    StructuredDocumentTagRow,
    StructuredDocumentTagRun,
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
                    children: [
                        new TextRun({ text: "SDT Content Controls Demo", bold: true, size: 32 }),
                    ],
                    spacing: { after: 400 },
                }),

                // Plain text SDT
                new Paragraph({
                    children: [new TextRun({ bold: true, text: "1. Plain Text SDT", size: 28 })],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new TextRun("Name: "),
                        new StructuredDocumentTagRun({
                            properties: {
                                alias: "Name",
                                text: { multiLine: false },
                            },
                            children: [new TextRun("John Doe")],
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // Multi-line text SDT
                new Paragraph({
                    children: [
                        new TextRun({ bold: true, text: "2. Multi-line Text SDT", size: 28 }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new StructuredDocumentTagRun({
                            properties: {
                                alias: "Description",
                                tag: "multiline-text",
                                text: { multiLine: true },
                            },
                            children: [new TextRun("This is a multi-line text content control.")],
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // ComboBox SDT
                new Paragraph({
                    children: [new TextRun({ bold: true, text: "3. ComboBox SDT", size: 28 })],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new TextRun("Color: "),
                        new StructuredDocumentTagRun({
                            properties: {
                                alias: "Color",
                                comboBox: {
                                    items: [
                                        { displayText: "Red", value: "red" },
                                        { displayText: "Blue", value: "blue" },
                                        { displayText: "Green", value: "green" },
                                    ],
                                    lastValue: "Red",
                                },
                            },
                            children: [new TextRun("Red")],
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // DropDownList SDT
                new Paragraph({
                    children: [new TextRun({ bold: true, text: "4. DropDownList SDT", size: 28 })],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new TextRun("Priority: "),
                        new StructuredDocumentTagRun({
                            properties: {
                                alias: "Priority",
                                dropDownList: {
                                    items: [
                                        { displayText: "High" },
                                        { displayText: "Medium" },
                                        { displayText: "Low" },
                                    ],
                                    lastValue: "Medium",
                                },
                            },
                            children: [new TextRun("Medium")],
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // Date SDT
                new Paragraph({
                    children: [new TextRun({ bold: true, text: "5. Date SDT", size: 28 })],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new TextRun("Date: "),
                        new StructuredDocumentTagRun({
                            properties: {
                                alias: "DueDate",
                                date: {
                                    dateFormat: "yyyy-MM-dd",
                                    languageId: "en-US",
                                    fullDate: "2026-04-29T00:00:00",
                                },
                            },
                            children: [new TextRun("2026-04-29")],
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // Block-level SDT (section child)
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "6. Block-level SDT (section child)",
                            size: 28,
                        }),
                    ],
                    spacing: { after: 200 },
                }),
                new StructuredDocumentTagBlock({
                    properties: {
                        richText: true,
                        alias: "BlockContent",
                        tag: "block-content",
                    },
                    children: [
                        new Paragraph({
                            children: [new TextRun("This is a block-level content control.")],
                        }),
                        new Paragraph({
                            children: [
                                new TextRun("It can contain multiple paragraphs and tables."),
                            ],
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // Cell-level SDT (wrapped table cell)
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "7. Cell-level SDT (wrapped table cell)",
                            size: 28,
                        }),
                    ],
                    spacing: { after: 200 },
                }),
                new Table({
                    columnWidths: [3000, 3000],
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [new Paragraph("Normal cell")],
                                }),
                                new StructuredDocumentTagCell({
                                    properties: {
                                        alias: "Controlled Cell",
                                        tag: "cell-sdt",
                                    },
                                    children: [
                                        new TableCell({
                                            children: [
                                                new Paragraph(
                                                    "This cell is wrapped in a content control.",
                                                ),
                                            ],
                                        }),
                                    ],
                                }),
                            ],
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // Row-level SDT (wrapped table row)
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "8. Row-level SDT (wrapped table row)",
                            size: 28,
                        }),
                    ],
                    spacing: { after: 200 },
                }),
                new Table({
                    columnWidths: [3000, 3000],
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [new Paragraph("Normal row - cell 1")],
                                }),
                                new TableCell({
                                    children: [new Paragraph("Normal row - cell 2")],
                                }),
                            ],
                        }),
                        new StructuredDocumentTagRow({
                            properties: {
                                alias: "Controlled Row",
                                tag: "row-sdt",
                            },
                            children: [
                                new TableRow({
                                    children: [
                                        new TableCell({
                                            children: [new Paragraph("Controlled row - cell 1")],
                                        }),
                                        new TableCell({
                                            children: [new Paragraph("Controlled row - cell 2")],
                                        }),
                                    ],
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
