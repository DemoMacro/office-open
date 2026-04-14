// The demo on the README.md

import * as fs from "fs";

import {
    Document,
    HeadingLevel,
    ImageRun,
    Packer,
    Paragraph,
    Table,
    TableCell,
    TableRow,
    VerticalAlignTable,
} from "docx-plus";

const table = new Table({
    rows: [
        new TableRow({
            children: [
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
                    verticalAlign: VerticalAlignTable.CENTER,
                }),
                new TableCell({
                    children: [
                        new Paragraph({
                            heading: HeadingLevel.HEADING_1,
                            text: "Hello",
                        }),
                    ],
                    verticalAlign: VerticalAlignTable.CENTER,
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [
                        new Paragraph({
                            heading: HeadingLevel.HEADING_1,
                            text: "World",
                        }),
                    ],
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
    ],
});

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    text: "Hello World",
                }),
                table,
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("./demo/images/pizza.gif"),
                            transformation: {
                                height: 100,
                                width: 100,
                            },
                            type: "gif",
                        }),
                    ],
                }),
            ],
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
