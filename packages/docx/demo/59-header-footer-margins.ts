// Move + offset header and footer

import * as fs from "fs";

import { Document, Footer, Header, Packer, Paragraph } from "docx-plus";

const doc = new Document({
    sections: [
        {
            children: [new Paragraph("Hello World")],
            footers: {
                default: new Footer({
                    children: [
                        new Paragraph({
                            indent: {
                                left: -400,
                            },
                            text: "Footer text",
                        }),
                    ],
                }),
            },
            headers: {
                default: new Header({
                    children: [
                        new Paragraph({
                            indent: {
                                left: -400,
                            },
                            text: "Header text",
                        }),
                        new Paragraph({
                            indent: {
                                left: -600,
                            },
                            text: "Some more header text",
                        }),
                    ],
                }),
            },
            properties: {
                page: {
                    margin: {
                        footer: 50,
                        header: 100,
                    },
                },
            },
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
