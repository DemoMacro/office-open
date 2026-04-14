// Add images to header and footer

import * as fs from "fs";

import { Document, Footer, Header, ImageRun, Packer, Paragraph } from "docx";

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
                                }),
                            ],
                        }),
                    ],
                }),
            },
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
