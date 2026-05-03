import * as fs from "fs";

import { Presentation, Packer, Shape, Paragraph, Run } from "@office-open/pptx";

const pres = new Presentation({
    title: "Rich Text Demo",
    creator: "Demo",
    slides: [
        {
            children: [
                // Title
                new Shape({
                    x: 50,
                    y: 30,
                    width: 600,
                    height: 50,
                    paragraphs: [
                        new Paragraph({
                            properties: { alignment: "ctr", bulletNone: true },
                            children: [
                                new Run({
                                    text: "Rich Text Formatting",
                                    fontSize: 36,
                                    bold: true,
                                    font: "Calibri",
                                }),
                            ],
                        }),
                    ],
                }),
                // Bold + Italic + Underline
                new Shape({
                    x: 50,
                    y: 100,
                    width: 600,
                    height: 40,
                    paragraphs: [
                        new Paragraph({
                            properties: { bulletNone: true },
                            children: [
                                new Run({ text: "Bold", bold: true, fontSize: 18 }),
                                new Run({ text: " | " }),
                                new Run({ text: "Italic", italic: true, fontSize: 18 }),
                                new Run({ text: " | " }),
                                new Run({ text: "Underline", underline: "sng", fontSize: 18 }),
                                new Run({ text: " | " }),
                                new Run({
                                    text: "Bold+Italic+Underline",
                                    bold: true,
                                    italic: true,
                                    underline: "dbl",
                                    fontSize: 18,
                                }),
                            ],
                        }),
                    ],
                }),
                // Strikethrough
                new Shape({
                    x: 50,
                    y: 150,
                    width: 600,
                    height: 40,
                    paragraphs: [
                        new Paragraph({
                            properties: { bulletNone: true },
                            children: [
                                new Run({
                                    text: "Single Strike",
                                    strike: "sngStrike",
                                    fontSize: 18,
                                }),
                                new Run({ text: " | " }),
                                new Run({
                                    text: "Double Strike",
                                    strike: "dblStrike",
                                    fontSize: 18,
                                }),
                            ],
                        }),
                    ],
                }),
                // Superscript + Subscript
                new Shape({
                    x: 50,
                    y: 200,
                    width: 600,
                    height: 40,
                    paragraphs: [
                        new Paragraph({
                            properties: { bulletNone: true },
                            children: [
                                new Run({ text: "E = mc", fontSize: 20 }),
                                new Run({ text: "2", fontSize: 14, baseline: 30000 }),
                                new Run({ text: "    H", fontSize: 20 }),
                                new Run({ text: "2", fontSize: 14, baseline: -25000 }),
                                new Run({ text: "O    x", fontSize: 20 }),
                                new Run({ text: "n+1", fontSize: 14, baseline: 30000 }),
                            ],
                        }),
                    ],
                }),
                // Character spacing
                new Shape({
                    x: 50,
                    y: 250,
                    width: 600,
                    height: 40,
                    paragraphs: [
                        new Paragraph({
                            properties: { bulletNone: true },
                            children: [
                                new Run({ text: "Normal spacing", fontSize: 18 }),
                                new Run({ text: "    Wide spacing", fontSize: 18, spacing: 400 }),
                                new Run({ text: "    Tight spacing", fontSize: 18, spacing: -100 }),
                            ],
                        }),
                    ],
                }),
                // Capitalization
                new Shape({
                    x: 50,
                    y: 300,
                    width: 600,
                    height: 40,
                    paragraphs: [
                        new Paragraph({
                            properties: { bulletNone: true },
                            children: [
                                new Run({ text: "Normal text", fontSize: 18 }),
                                new Run({
                                    text: "    ALL CAPS",
                                    fontSize: 18,
                                    capitalization: "all",
                                }),
                                new Run({
                                    text: "    Small Caps",
                                    fontSize: 18,
                                    capitalization: "small",
                                }),
                            ],
                        }),
                    ],
                }),
                // Text color
                new Shape({
                    x: 50,
                    y: 350,
                    width: 600,
                    height: 40,
                    fill: "333333",
                    paragraphs: [
                        new Paragraph({
                            properties: { bulletNone: true },
                            children: [
                                new Run({
                                    text: "Red",
                                    fontSize: 20,
                                    fill: "FF0000",
                                }),
                                new Run({ text: " | " }),
                                new Run({
                                    text: "Green",
                                    fontSize: 20,
                                    fill: "00FF00",
                                }),
                                new Run({ text: " | " }),
                                new Run({
                                    text: "Blue",
                                    fontSize: 20,
                                    fill: "4472C4",
                                }),
                                new Run({ text: " | " }),
                                new Run({
                                    text: "Yellow",
                                    fontSize: 20,
                                    fill: "FFC000",
                                }),
                            ],
                        }),
                    ],
                }),
                // Alignment demo
                new Shape({
                    x: 50,
                    y: 420,
                    width: 600,
                    height: 120,
                    outline: { color: "999999", width: 1 },
                    paragraphs: [
                        new Paragraph({
                            properties: { alignment: "l", bulletNone: true },
                            children: [new Run({ text: "Left aligned text", fontSize: 16 })],
                        }),
                        new Paragraph({
                            properties: { alignment: "ctr", bulletNone: true },
                            children: [new Run({ text: "Center aligned text", fontSize: 16 })],
                        }),
                        new Paragraph({
                            properties: { alignment: "r", bulletNone: true },
                            children: [new Run({ text: "Right aligned text", fontSize: 16 })],
                        }),
                        new Paragraph({
                            properties: { alignment: "just", bulletNone: true },
                            children: [
                                new Run({
                                    text: "Justified text that is long enough to wrap to multiple lines to demonstrate the justification effect clearly",
                                    fontSize: 16,
                                }),
                            ],
                        }),
                    ],
                }),
                // Bullet & numbering
                new Shape({
                    x: 50,
                    y: 710,
                    width: 280,
                    height: 200,
                    outline: { color: "4472C4", width: 1 },
                    paragraphs: [
                        new Paragraph({
                            properties: { bullet: { type: "char", char: "●" } },
                            children: [new Run({ text: "Bullet Point 1" })],
                        }),
                        new Paragraph({
                            properties: { bullet: { type: "char", char: "●", color: "4472C4" } },
                            children: [new Run({ text: "Blue Bullet 2" })],
                        }),
                        new Paragraph({
                            properties: { bullet: { type: "char", char: "■", size: 75 } },
                            children: [new Run({ text: "Small Square 3" })],
                        }),
                        new Paragraph({
                            properties: { bullet: { type: "char", char: "➢" } },
                            children: [new Run({ text: "Arrow Bullet 4" })],
                        }),
                    ],
                }),
                new Shape({
                    x: 360,
                    y: 710,
                    width: 280,
                    height: 200,
                    outline: { color: "ED7D31", width: 1 },
                    paragraphs: [
                        new Paragraph({
                            properties: { bullet: { type: "autoNum", format: "arabicPeriod" } },
                            children: [new Run({ text: "First item" })],
                        }),
                        new Paragraph({
                            properties: { bullet: { type: "autoNum", format: "arabicPeriod" } },
                            children: [new Run({ text: "Second item" })],
                        }),
                        new Paragraph({
                            properties: { bullet: { type: "autoNum", format: "arabicPeriod" } },
                            children: [new Run({ text: "Third item" })],
                        }),
                        new Paragraph({
                            properties: { bullet: { type: "autoNum", format: "alphaLcParenBoth" } },
                            children: [new Run({ text: "Sub-item a" })],
                        }),
                        new Paragraph({
                            properties: { bullet: { type: "autoNum", format: "alphaLcParenBoth" } },
                            children: [new Run({ text: "Sub-item b" })],
                        }),
                    ],
                }),
                // RTL, NoProof, Shadow, Outline
                new Shape({
                    x: 660,
                    y: 710,
                    width: 280,
                    height: 200,
                    outline: { color: "70AD47", width: 1 },
                    paragraphs: [
                        new Paragraph({
                            children: [new Run({ text: "Normal text", fontSize: 16 })],
                        }),
                        new Paragraph({
                            children: [
                                new Run({ text: "Right-to-Left", rightToLeft: true, fontSize: 16 }),
                            ],
                        }),
                        new Paragraph({
                            children: [
                                new Run({ text: "No proof text", noProof: true, fontSize: 16 }),
                            ],
                        }),
                        new Paragraph({
                            children: [
                                new Run({ text: "Shadow text", shadow: true, fontSize: 16 }),
                            ],
                        }),
                        new Paragraph({
                            children: [
                                new Run({ text: "Outline text", outline: true, fontSize: 16 }),
                            ],
                        }),
                    ],
                }),
                // Line spacing demo
                new Shape({
                    x: 50,
                    y: 560,
                    width: 280,
                    height: 120,
                    outline: { color: "4472C4", width: 1 },
                    paragraphs: [
                        new Paragraph({
                            properties: { lineSpacing: 100, bulletNone: true },
                            children: [new Run({ text: "Single spacing (1.0)", fontSize: 14 })],
                        }),
                        new Paragraph({
                            properties: { lineSpacing: 150, bulletNone: true },
                            children: [new Run({ text: "1.5x spacing", fontSize: 14 })],
                        }),
                        new Paragraph({
                            properties: { lineSpacing: 200, bulletNone: true },
                            children: [new Run({ text: "Double spacing (2.0)", fontSize: 14 })],
                        }),
                    ],
                }),
                // Fixed line spacing
                new Shape({
                    x: 360,
                    y: 560,
                    width: 280,
                    height: 120,
                    outline: { color: "ED7D31", width: 1 },
                    paragraphs: [
                        new Paragraph({
                            properties: { lineSpacingPoints: 14, bulletNone: true },
                            children: [new Run({ text: "Exactly 14pt", fontSize: 14 })],
                        }),
                        new Paragraph({
                            properties: { lineSpacingPoints: 20, bulletNone: true },
                            children: [new Run({ text: "Exactly 20pt", fontSize: 14 })],
                        }),
                        new Paragraph({
                            properties: { lineSpacingPoints: 28, bulletNone: true },
                            children: [new Run({ text: "Exactly 28pt", fontSize: 14 })],
                        }),
                    ],
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(pres);
fs.writeFileSync("My Presentation.pptx", buffer);
