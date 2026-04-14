import * as fs from "fs";

// Simple example to add textbox to a document
import { Document, Packer, Paragraph, TextRun, Textbox } from "docx";

const doc = new Document({
    sections: [
        {
            children: [
                new Textbox({
                    alignment: "center",
                    children: [
                        new Paragraph({
                            children: [new TextRun("Hi i'm a textbox!")],
                        }),
                    ],
                    style: {
                        height: "auto",
                        width: "200pt",
                    },
                }),
                new Textbox({
                    alignment: "center",
                    children: [
                        new Paragraph({
                            children: [new TextRun("Hi i'm a textbox with a hidden box!")],
                        }),
                    ],
                    style: {
                        height: 400,
                        visibility: "hidden",
                        width: "300pt",
                        zIndex: "auto",
                    },
                }),
            ],
            properties: {},
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
