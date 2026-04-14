// Scaling images

import * as fs from "fs";

import { Document, ImageRun, Packer, Paragraph } from "docx";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph("Hello World"),
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("./demo/images/pizza.gif"),
                            transformation: {
                                height: 50,
                                width: 50,
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
                            data: fs.readFileSync("./demo/images/pizza.gif"),
                            transformation: {
                                height: 250,
                                width: 250,
                            },
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("./demo/images/pizza.gif"),
                            transformation: {
                                height: 400,
                                width: 400,
                            },
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
