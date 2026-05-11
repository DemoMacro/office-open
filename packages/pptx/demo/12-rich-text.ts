import * as fs from "fs";

import { Presentation, Packer, Shape, Paragraph, TextRun } from "@office-open/pptx";

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
                            properties: { alignment: "CENTER", bulletNone: true },
                            children: [
                                new TextRun({
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
                                new TextRun({ text: "Bold", bold: true, fontSize: 18 }),
                                new TextRun({ text: " | " }),
                                new TextRun({ text: "Italic", italic: true, fontSize: 18 }),
                                new TextRun({ text: " | " }),
                                new TextRun({
                                    text: "Underline",
                                    underline: "SINGLE",
                                    fontSize: 18,
                                }),
                                new TextRun({ text: " | " }),
                                new TextRun({
                                    text: "Bold+Italic+Underline",
                                    bold: true,
                                    italic: true,
                                    underline: "DOUBLE",
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
                                new TextRun({
                                    text: "Single Strike",
                                    strike: "SINGLE",
                                    fontSize: 18,
                                }),
                                new TextRun({ text: " | " }),
                                new TextRun({
                                    text: "Double Strike",
                                    strike: "DOUBLE",
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
                                new TextRun({ text: "E = mc", fontSize: 20 }),
                                new TextRun({ text: "2", fontSize: 14, baseline: 30000 }),
                                new TextRun({ text: "    H", fontSize: 20 }),
                                new TextRun({ text: "2", fontSize: 14, baseline: -25000 }),
                                new TextRun({ text: "O    x", fontSize: 20 }),
                                new TextRun({ text: "n+1", fontSize: 14, baseline: 30000 }),
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
                                new TextRun({ text: "Normal spacing", fontSize: 18 }),
                                new TextRun({
                                    text: "    Wide spacing",
                                    fontSize: 18,
                                    spacing: 400,
                                }),
                                new TextRun({
                                    text: "    Tight spacing",
                                    fontSize: 18,
                                    spacing: -100,
                                }),
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
                                new TextRun({ text: "Normal text", fontSize: 18 }),
                                new TextRun({
                                    text: "    ALL CAPS",
                                    fontSize: 18,
                                    capitalization: "ALL",
                                }),
                                new TextRun({
                                    text: "    Small Caps",
                                    fontSize: 18,
                                    capitalization: "SMALL",
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
                                new TextRun({
                                    text: "Red",
                                    fontSize: 20,
                                    fill: "FF0000",
                                }),
                                new TextRun({ text: " | " }),
                                new TextRun({
                                    text: "Green",
                                    fontSize: 20,
                                    fill: "00FF00",
                                }),
                                new TextRun({ text: " | " }),
                                new TextRun({
                                    text: "Blue",
                                    fontSize: 20,
                                    fill: "4472C4",
                                }),
                                new TextRun({ text: " | " }),
                                new TextRun({
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
                            properties: { alignment: "LEFT", bulletNone: true },
                            children: [new TextRun({ text: "Left aligned text", fontSize: 16 })],
                        }),
                        new Paragraph({
                            properties: { alignment: "CENTER", bulletNone: true },
                            children: [new TextRun({ text: "Center aligned text", fontSize: 16 })],
                        }),
                        new Paragraph({
                            properties: { alignment: "RIGHT", bulletNone: true },
                            children: [new TextRun({ text: "Right aligned text", fontSize: 16 })],
                        }),
                        new Paragraph({
                            properties: { alignment: "JUSTIFY", bulletNone: true },
                            children: [
                                new TextRun({
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
                            children: [new TextRun({ text: "Bullet Point 1" })],
                        }),
                        new Paragraph({
                            properties: { bullet: { type: "char", char: "●", color: "4472C4" } },
                            children: [new TextRun({ text: "Blue Bullet 2" })],
                        }),
                        new Paragraph({
                            properties: { bullet: { type: "char", char: "■", size: 75 } },
                            children: [new TextRun({ text: "Small Square 3" })],
                        }),
                        new Paragraph({
                            properties: { bullet: { type: "char", char: "➢" } },
                            children: [new TextRun({ text: "Arrow Bullet 4" })],
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
                            children: [new TextRun({ text: "First item" })],
                        }),
                        new Paragraph({
                            properties: { bullet: { type: "autoNum", format: "arabicPeriod" } },
                            children: [new TextRun({ text: "Second item" })],
                        }),
                        new Paragraph({
                            properties: { bullet: { type: "autoNum", format: "arabicPeriod" } },
                            children: [new TextRun({ text: "Third item" })],
                        }),
                        new Paragraph({
                            properties: { bullet: { type: "autoNum", format: "alphaLcParenBoth" } },
                            children: [new TextRun({ text: "Sub-item a" })],
                        }),
                        new Paragraph({
                            properties: { bullet: { type: "autoNum", format: "alphaLcParenBoth" } },
                            children: [new TextRun({ text: "Sub-item b" })],
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
                            children: [new TextRun({ text: "Normal text", fontSize: 16 })],
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "Right-to-Left",
                                    rightToLeft: true,
                                    fontSize: 16,
                                }),
                            ],
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({ text: "No proof text", noProof: true, fontSize: 16 }),
                            ],
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({ text: "Shadow text", shadow: true, fontSize: 16 }),
                            ],
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({ text: "Outline text", outline: true, fontSize: 16 }),
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
                            children: [new TextRun({ text: "Single spacing (1.0)", fontSize: 14 })],
                        }),
                        new Paragraph({
                            properties: { lineSpacing: 150, bulletNone: true },
                            children: [new TextRun({ text: "1.5x spacing", fontSize: 14 })],
                        }),
                        new Paragraph({
                            properties: { lineSpacing: 200, bulletNone: true },
                            children: [new TextRun({ text: "Double spacing (2.0)", fontSize: 14 })],
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
                            children: [new TextRun({ text: "Exactly 14pt", fontSize: 14 })],
                        }),
                        new Paragraph({
                            properties: { lineSpacingPoints: 20, bulletNone: true },
                            children: [new TextRun({ text: "Exactly 20pt", fontSize: 14 })],
                        }),
                        new Paragraph({
                            properties: { lineSpacingPoints: 28, bulletNone: true },
                            children: [new TextRun({ text: "Exactly 28pt", fontSize: 14 })],
                        }),
                    ],
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(pres);
fs.writeFileSync("My Presentation.pptx", buffer);
