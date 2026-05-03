// Example of how to set the document to landscape

import * as fs from "fs";

import { Document, Packer, PageOrientation, Paragraph } from "@office-open/docx";

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

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
