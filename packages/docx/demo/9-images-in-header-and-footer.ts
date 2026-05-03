// Add images to header and footer

import * as fs from "fs";

import { Document, Footer, Header, ImageRun, Packer, Paragraph } from "@office-open/docx";

const doc = new Document({
    sections: [
        {
            children: [new Paragraph("Hello World")],
            footers: {
                default: new Footer({
                    children: [
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
                }),
            },
            headers: {
                default: new Header({
                    children: [
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
                }),
            },
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
