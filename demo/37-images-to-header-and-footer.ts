// Add images to header and footer

import * as fs from "fs";

import { Document, Header, ImageRun, Packer, Paragraph } from "docx-plus";

const doc = new Document({
    sections: [
        {
            children: [new Paragraph("Hello World")],
            headers: {
                default: new Header({
                    children: [
                        new Paragraph({
                            children: [
                                new ImageRun({
                                    data: fs.readFileSync("./demo/images/image1.jpeg"),
                                    transformation: {
                                        height: 100,
                                        width: 100,
                                    },
                                }),
                            ],
                        }),
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
                        new Paragraph({
                            children: [
                                new ImageRun({
                                    data: fs.readFileSync("./demo/images/image1.jpeg"),
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
