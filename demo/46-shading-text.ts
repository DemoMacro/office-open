// Shading text

import * as fs from "fs";

import {
    AlignmentType,
    Document,
    Header,
    Packer,
    Paragraph,
    ShadingType,
    TextRun,
} from "docx-plus";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun({
                            emboss: true,
                            text: "Embossed text - hello world",
                        }),
                        new TextRun({
                            imprint: true,
                            text: "Imprinted text - hello world",
                        }),
                    ],
                }),
            ],
            headers: {
                default: new Header({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.RIGHT,
                            children: [
                                new TextRun({
                                    bold: true,
                                    color: "FF0000",
                                    font: {
                                        name: "Garamond",
                                    },
                                    shading: {
                                        color: "00FFFF",
                                        fill: "FF0000",
                                        type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                                    },
                                    size: 24,
                                    text: "Hello World",
                                }),
                            ],
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "Hello World for entire paragraph",
                                }),
                            ],
                            shading: {
                                color: "00FFFF",
                                fill: "FF0000",
                                type: ShadingType.DIAGONAL_CROSS,
                            },
                        }),
                    ],
                }),
            },
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
