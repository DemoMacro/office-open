// Simple example to add text to a document

import * as fs from "fs";

import {
    Document,
    Math,
    MathAngledBrackets,
    MathCurlyBrackets,
    MathFraction,
    MathFunction,
    MathIntegral,
    MathLimitLower,
    MathLimitUpper,
    MathPreSubSuperScript,
    MathRadical,
    MathRoundBrackets,
    MathRun,
    MathSquareBrackets,
    MathSubScript,
    MathSubSuperScript,
    MathSum,
    MathSuperScript,
    Packer,
    Paragraph,
    TextRun,
} from "docx-plus";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathRun("2+2"),
                                new MathFraction({
                                    denominator: [new MathRun("2")],
                                    numerator: [new MathRun("hi")],
                                }),
                            ],
                        }),
                        new TextRun({
                            bold: true,
                            text: "Foo Bar",
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathFraction({
                                    denominator: [new MathRun("2")],
                                    numerator: [
                                        new MathRun("1"),
                                        new MathRadical({
                                            children: [new MathRun("2")],
                                        }),
                                    ],
                                }),
                            ],
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathSum({
                                    children: [new MathRun("test")],
                                }),
                                new MathSum({
                                    children: [
                                        new MathSuperScript({
                                            children: [new MathRun("e")],
                                            superScript: [new MathRun("2")],
                                        }),
                                    ],
                                    subScript: [new MathRun("i")],
                                }),
                                new MathSum({
                                    children: [
                                        new MathRadical({
                                            children: [new MathRun("i")],
                                        }),
                                    ],
                                    subScript: [new MathRun("i")],
                                    superScript: [new MathRun("10")],
                                }),
                            ],
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathIntegral({
                                    children: [new MathRun("test")],
                                }),
                                new MathIntegral({
                                    children: [
                                        new MathSuperScript({
                                            children: [new MathRun("e")],
                                            superScript: [new MathRun("2")],
                                        }),
                                    ],
                                    subScript: [new MathRun("i")],
                                }),
                                new MathIntegral({
                                    children: [
                                        new MathRadical({
                                            children: [new MathRun("i")],
                                        }),
                                    ],
                                    subScript: [new MathRun("i")],
                                    superScript: [new MathRun("10")],
                                }),
                            ],
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathSuperScript({
                                    children: [new MathRun("test")],
                                    superScript: [new MathRun("hello")],
                                }),
                            ],
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathSubScript({
                                    children: [new MathRun("test")],
                                    subScript: [new MathRun("hello")],
                                }),
                            ],
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathSubScript({
                                    children: [new MathRun("x")],
                                    subScript: [
                                        new MathSuperScript({
                                            children: [new MathRun("y")],
                                            superScript: [new MathRun("2")],
                                        }),
                                    ],
                                }),
                            ],
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathSubSuperScript({
                                    children: [new MathRun("test")],
                                    subScript: [new MathRun("world")],
                                    superScript: [new MathRun("hello")],
                                }),
                            ],
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathPreSubSuperScript({
                                    children: [new MathRun("test")],
                                    subScript: [new MathRun("world")],
                                    superScript: [new MathRun("hello")],
                                }),
                            ],
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathSubScript({
                                    children: [
                                        new MathFraction({
                                            denominator: [new MathRun("2")],
                                            numerator: [new MathRun("1")],
                                        }),
                                    ],
                                    subScript: [new MathRun("4")],
                                }),
                            ],
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathSubScript({
                                    children: [
                                        new MathRadical({
                                            children: [
                                                new MathFraction({
                                                    denominator: [new MathRun("2")],
                                                    numerator: [new MathRun("1")],
                                                }),
                                            ],
                                            degree: [new MathRun("4")],
                                        }),
                                    ],
                                    subScript: [new MathRun("x")],
                                }),
                            ],
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathRadical({
                                    children: [new MathRun("4")],
                                }),
                            ],
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathFunction({
                                    children: [new MathRun("100")],
                                    name: [
                                        new MathSuperScript({
                                            children: [new MathRun("cos")],
                                            superScript: [new MathRun("-1")],
                                        }),
                                    ],
                                }),
                                new MathRun("×"),
                                new MathFunction({
                                    children: [new MathRun("360")],
                                    name: [new MathRun("sin")],
                                }),
                                new MathRun("= x"),
                            ],
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathRoundBrackets({
                                    children: [
                                        new MathFraction({
                                            denominator: [new MathRun("2")],
                                            numerator: [new MathRun("1")],
                                        }),
                                    ],
                                }),
                                new MathSquareBrackets({
                                    children: [
                                        new MathFraction({
                                            denominator: [new MathRun("2")],
                                            numerator: [new MathRun("1")],
                                        }),
                                    ],
                                }),
                                new MathCurlyBrackets({
                                    children: [
                                        new MathFraction({
                                            denominator: [new MathRun("2")],
                                            numerator: [new MathRun("1")],
                                        }),
                                    ],
                                }),
                                new MathAngledBrackets({
                                    children: [
                                        new MathFraction({
                                            denominator: [new MathRun("2")],
                                            numerator: [new MathRun("1")],
                                        }),
                                    ],
                                }),
                            ],
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathFraction({
                                    denominator: [new MathRun("2a")],
                                    numerator: [
                                        new MathRadical({
                                            children: [new MathRun("4")],
                                        }),
                                    ],
                                }),
                            ],
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathLimitUpper({
                                    children: [new MathRun("x")],
                                    limit: [new MathRun("-")],
                                }),
                                new MathRun("="),
                                new MathLimitLower({
                                    children: [new MathRun("lim")],
                                    limit: [new MathRun("x→0")],
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
