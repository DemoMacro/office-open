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

const styles = fs.readFileSync("./demo/assets/custom-styles.xml", "utf8");

// Create a table and pass the XML Style
const table = new Table({
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("Header Colum 1")],
                }),
                new TableCell({
                    children: [new Paragraph("Header Colum 2")],
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("Column Content 3")],
                }),
                new TableCell({
                    children: [new Paragraph("Column Content 2")],
                }),
            ],
        }),
    ],
    style: "MyCustomTableStyle",
    width: {
        size: 9070,
        type: WidthType.DXA,
    },
});

const doc = new Document({
    externalStyles: styles,
    sections: [
        {
            children: [table],
        },
    ],
    title: "Title",
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
