// East Asian layout - Need to use an East Asian font
import * as fs from "fs";

import { Document, HeadingLevel, Packer, Paragraph, TextRun } from "docx-plus";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    text: "East Asian Layout",
                }),

                // East Asian layout - combined characters with round brackets
                new Paragraph({
                    children: [
                        new TextRun("Combined characters (round brackets): "),
                        new TextRun({
                            eastAsianLayout: {
                                combine: true,
                                combineBrackets: "round",
                                id: 1,
                            },
                            text: "国民",
                        }),
                    ],
                    spacing: { after: 200 },
                }),

                // East Asian layout - combined characters with square brackets
                new Paragraph({
                    children: [
                        new TextRun("Combined characters (square brackets): "),
                        new TextRun({
                            eastAsianLayout: {
                                combine: true,
                                combineBrackets: "square",
                                id: 2,
                            },
                            text: "日本語",
                        }),
                    ],
                    spacing: { after: 200 },
                }),

                // East Asian layout - vertical text
                new Paragraph({
                    children: [
                        new TextRun("Vertical text: "),
                        new TextRun({
                            eastAsianLayout: {
                                vert: true,
                            },
                            text: "縦書き",
                        }),
                    ],
                    spacing: { after: 200 },
                }),
            ],
        },
    ],
    styles: {
        paragraphStyles: [
            {
                basedOn: "Normal",
                id: "Normal",
                name: "Normal",
                next: "Normal",
                quickFormat: true,
                run: {
                    font: "MS Gothic",
                },
            },
        ],
    },
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
