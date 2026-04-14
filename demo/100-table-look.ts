import * as fs from "fs";

// Example of using tableLook to control conditional table formatting
import { Document, Packer, Paragraph, Table, TableCell, TableRow, TextRun, WidthType } from "docx";

const styles = fs.readFileSync("./demo/assets/custom-styles.xml", "utf8");

const doc = new Document({
    externalStyles: styles,
    sections: [
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun({ bold: true, text: "Table 1: Table Look Default Values" }),
                    ],
                }),
                new Paragraph({ text: "" }),
                new Table({
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph("Header 1")] }),
                                new TableCell({ children: [new Paragraph("Header 2")] }),
                                new TableCell({ children: [new Paragraph("Header 3")] }),
                            ],
                        }),
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph("Row 1, Col 1")] }),
                                new TableCell({ children: [new Paragraph("Row 1, Col 2")] }),
                                new TableCell({ children: [new Paragraph("Row 1, Col 3")] }),
                            ],
                        }),
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph("Row 2, Col 1")] }),
                                new TableCell({ children: [new Paragraph("Row 2, Col 2")] }),
                                new TableCell({ children: [new Paragraph("Row 2, Col 3")] }),
                            ],
                        }),
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph("Row 3, Col 1")] }),
                                new TableCell({ children: [new Paragraph("Row 3, Col 2")] }),
                                new TableCell({ children: [new Paragraph("Row 3, Col 3")] }),
                            ],
                        }),
                    ],
                    style: "MyCustomTableStyle",
                    width: {
                        size: 100,
                        type: WidthType.PERCENTAGE,
                    },
                }),
                new Paragraph({ text: "" }),
                new Paragraph({ text: "" }),
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "Table 2: Table Look All Look Values Enabled",
                        }),
                    ],
                }),
                new Paragraph({ text: "" }),
                new Table({
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph("Header 1")] }),
                                new TableCell({ children: [new Paragraph("Header 2")] }),
                                new TableCell({ children: [new Paragraph("Header 3")] }),
                            ],
                        }),
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph("Row 1, Col 1")] }),
                                new TableCell({ children: [new Paragraph("Row 1, Col 2")] }),
                                new TableCell({ children: [new Paragraph("Row 1, Col 3")] }),
                            ],
                        }),
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph("Row 2, Col 1")] }),
                                new TableCell({ children: [new Paragraph("Row 2, Col 2")] }),
                                new TableCell({ children: [new Paragraph("Row 2, Col 3")] }),
                            ],
                        }),
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph("Row 3, Col 1")] }),
                                new TableCell({ children: [new Paragraph("Row 3, Col 2")] }),
                                new TableCell({ children: [new Paragraph("Row 3, Col 3")] }),
                            ],
                        }),
                    ],
                    style: "MyCustomTableStyle",
                    tableLook: {
                        firstColumn: true,
                        firstRow: true,
                        lastColumn: true,
                        lastRow: true,
                        noHBand: false,
                        noVBand: false,
                    },
                    width: {
                        size: 100,
                        type: WidthType.PERCENTAGE,
                    },
                }),
            ],
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("97-table-look.docx", buffer);
    console.log("Document created successfully at 97-table-look.docx");
});
