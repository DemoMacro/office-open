// Custom character styles using JavaScript configuration

import * as fs from "fs";

import { Document, Packer, Paragraph, TextRun } from "@office-open/docx";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun({
                            style: "myRedStyle",
                            text: "Foo bar",
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            style: "strong",
                            text: "First Word",
                        }),
                        new TextRun({
                            text: " - Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.",
                        }),
                    ],
                }),
            ],
        },
    ],
    styles: {
        characterStyles: [
            {
                basedOn: "Normal",
                id: "myRedStyle",
                name: "My Wonky Style",
                run: {
                    color: "990000",
                    italics: true,
                },
            },
            {
                basedOn: "Normal",
                id: "strong",
                name: "Strong",
                run: {
                    bold: true,
                },
            },
        ],
    },
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
