// Simple example to add check boxes to a document
import * as fs from "fs";

import { CheckBox, Document, Packer, Paragraph, TextRun } from "docx-plus";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun("Hello World"),
                        new TextRun({ break: 1 }),
                        new CheckBox(),
                        new TextRun({ break: 1 }),
                        new CheckBox({ checked: true }),
                        new TextRun({ break: 1 }),
                        new CheckBox({ checked: true, checkedState: { value: "2611" } }),
                        new TextRun({ break: 1 }),
                        new CheckBox({
                            checked: true,
                            checkedState: { font: "MS Gothic", value: "2611" },
                        }),
                        new TextRun({ break: 1 }),
                        new CheckBox({
                            checked: true,
                            checkedState: { font: "MS Gothic", value: "2611" },
                            uncheckedState: { font: "MS Gothic", value: "2610" },
                        }),
                        new TextRun({ break: 1 }),
                        new CheckBox({
                            checked: true,
                            checkedState: { font: "MS Gothic", value: "2611" },
                            uncheckedState: { font: "MS Gothic", value: "2610" },
                        }),
                        new TextRun({ break: 1, text: "Are you ok?" }),
                        new CheckBox({ alias: "Are you ok?", checked: true }),
                    ],
                }),
            ],
            properties: {},
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
