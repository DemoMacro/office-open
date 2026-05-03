// Demo: New OOXML features - Math accent, Fraction types

import * as fs from "fs";

import {
    createMathAccent,
    Document,
    Math,
    MathFraction,
    MathRun,
    Packer,
    Paragraph,
    TextRun,
} from "@office-open/docx";

const doc = new Document({
    sections: [
        {
            children: [
                // ===== Math Accent =====
                new Paragraph({
                    children: [new TextRun({ text: "Math Accent Demo", bold: true, size: 28 })],
                }),
                new Paragraph({
                    children: [
                        new Math({
                            children: [
                                createMathAccent({
                                    children: [new MathRun("x")],
                                }),
                                new MathRun(" + "),
                                createMathAccent({
                                    accentCharacter: "\u0303",
                                    children: [new MathRun("y")],
                                }),
                                new MathRun(" + "),
                                createMathAccent({
                                    accentCharacter: "\u0307",
                                    children: [new MathRun("z")],
                                }),
                            ],
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new Math({
                            children: [
                                createMathAccent({
                                    accentCharacter: "\u20D7",
                                    children: [new MathRun("AB")],
                                }),
                                new MathRun(" = "),
                                createMathAccent({
                                    accentCharacter: "\u0305",
                                    children: [new MathRun("CD")],
                                }),
                            ],
                        }),
                    ],
                }),

                // ===== Fraction Types =====
                new Paragraph({
                    children: [new TextRun({ text: "Fraction Types Demo", bold: true, size: 28 })],
                }),
                new Paragraph({
                    children: [
                        new TextRun("Standard bar fraction: "),
                        new Math({
                            children: [
                                new MathFraction({
                                    numerator: [new MathRun("a")],
                                    denominator: [new MathRun("b")],
                                }),
                            ],
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun("Skewed fraction: "),
                        new Math({
                            children: [
                                new MathFraction({
                                    numerator: [new MathRun("a")],
                                    denominator: [new MathRun("b")],
                                    fractionType: "SKEWED",
                                }),
                            ],
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun("Linear fraction: "),
                        new Math({
                            children: [
                                new MathFraction({
                                    numerator: [new MathRun("a")],
                                    denominator: [new MathRun("b")],
                                    fractionType: "LINEAR",
                                }),
                            ],
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun("No-bar fraction: "),
                        new Math({
                            children: [
                                new MathFraction({
                                    numerator: [new MathRun("a")],
                                    denominator: [new MathRun("b")],
                                    fractionType: "NO_BAR",
                                }),
                            ],
                        }),
                    ],
                }),
            ],
            properties: {},
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
