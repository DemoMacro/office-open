// Export to base64 string - Useful in a browser environment.

import * as fs from "fs";

import { Document, Packer, Paragraph, Tab, TextRun } from "@office-open/docx";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun("Hello World"),
                        new TextRun({
                            bold: true,
                            text: "Foo",
                        }),
                        new TextRun({
                            bold: true,
                            children: [new Tab(), "Bar"],
                        }),
                    ],
                }),
            ],
        },
    ],
});

const str = await Packer.toBase64String(doc);
fs.writeFileSync("My Document.docx", str);
