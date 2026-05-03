// Add image to table cell

import * as fs from "fs";

import {
    Document,
    ImageRun,
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
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new ImageRun({
                                                    data: fs.readFileSync(
                                                        "./demo/images/image1.jpeg",
                                                    ),
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
                                    children: [new Paragraph("Hello")],
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
