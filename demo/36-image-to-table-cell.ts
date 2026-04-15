// Add image to table cell in a header and body

import * as fs from "fs";

import {
    Document,
    Header,
    ImageRun,
    Packer,
    Paragraph,
    Table,
    TableCell,
    TableRow,
} from "docx-plus";

const table = new Table({
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
                    children: [
                        new Paragraph({
                            children: [
                                new ImageRun({
                                    data: fs.readFileSync("./demo/images/image1.jpeg"),
                                    transformation: {
                                        height: 100,
                                        width: 100,
                                    },
                                    type: "jpg",
                                }),
                            ],
                        }),
                    ],
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
});

// Adding same table in the body and in the header
const doc = new Document({
    sections: [
        {
            children: [table],
            headers: {
                default: new Header({
                    children: [table],
                }),
            },
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
