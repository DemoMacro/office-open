// Section with 2 columns including a column break

import * as fs from "fs";

import { ColumnBreak, Document, Packer, Paragraph, TextRun } from "docx";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun("This text will be in the first column."),
                        new ColumnBreak(),
                        new TextRun("This text will be in the second column."),
                    ],
                }),
            ],
            properties: {
                column: {
                    count: 2,
                    space: 708,
                },
            },
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
