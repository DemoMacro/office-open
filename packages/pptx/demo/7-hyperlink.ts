import * as fs from "fs";

import { Presentation, Packer, Shape, SolidFill, Run, Paragraph } from "../src";

const pres = new Presentation({
    title: "Hyperlink Demo",
    creator: "Demo",
    slides: [
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 400,
                    height: 60,
                    text: "Hyperlinks in PPTX",
                    fill: new SolidFill("4472C4"),
                }),
                new Shape({
                    x: 50,
                    y: 120,
                    width: 600,
                    height: 200,
                    paragraphs: [
                        new Paragraph({
                            children: [
                                new Run({
                                    text: "Visit ",
                                }),
                                new Run({
                                    text: "Google",
                                    hyperlink: { url: "https://www.google.com", tooltip: "Go to Google" },
                                    bold: true,
                                }),
                                new Run({
                                    text: " or ",
                                }),
                                new Run({
                                    text: "GitHub",
                                    hyperlink: { url: "https://github.com" },
                                    italic: true,
                                }),
                            ],
                        }),
                        new Paragraph({
                            children: [
                                new Run({
                                    text: "Another link: ",
                                }),
                                new Run({
                                    text: "Office Open XML Spec",
                                    hyperlink: {
                                        url: "https://www.iso.org/standard/71691.html",
                                        tooltip: "ISO/IEC 29500",
                                    },
                                    fontSize: 14,
                                }),
                            ],
                        }),
                    ],
                    fill: new SolidFill("F2F2F2"),
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(pres);
fs.writeFileSync("My Presentation.pptx", buffer);
