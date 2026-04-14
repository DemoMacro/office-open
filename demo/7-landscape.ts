// Example of how to set the document to landscape

import * as fs from "fs";

import { Document, Packer, PageOrientation, Paragraph } from "docx";

const doc = new Document({
    sections: [
        {
            children: [new Paragraph("Hello World")],
            properties: {
                page: {
                    size: {
                        orientation: PageOrientation.LANDSCAPE,
                    },
                },
            },
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
