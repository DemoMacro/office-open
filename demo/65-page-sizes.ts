// Example of how to set the document page sizes
// Reference from https://papersizes.io/a/a3

import * as fs from "fs";

import { Document, Packer, PageOrientation, Paragraph, convertMillimetersToTwip } from "docx";

const doc = new Document({
    sections: [
        {
            children: [new Paragraph("Hello World")],
            properties: {
                page: {
                    size: {
                        height: convertMillimetersToTwip(210),
                        orientation: PageOrientation.LANDSCAPE,
                        width: convertMillimetersToTwip(148),
                    },
                },
            },
        },
        {
            children: [new Paragraph("Hello World")],
            properties: {
                page: {
                    size: {
                        height: convertMillimetersToTwip(420),
                        orientation: PageOrientation.PORTRAIT,
                        width: convertMillimetersToTwip(297),
                    },
                },
            },
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
