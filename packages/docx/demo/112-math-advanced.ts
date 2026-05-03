// Demo: Advanced math elements - box, borderBox, eqArr, groupChr, matrix, phantom
import * as fs from "fs";

import {
    Document,
    Math,
    MathBorderBox,
    MathBox,
    MathEqArr,
    MathFraction,
    MathGroupChr,
    MathMatrix,
    MathPhant,
    MathRun,
    MathSuperScript,
    Packer,
    Paragraph,
    TextRun,
} from "@office-open/docx";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun({ text: "Advanced Math Elements Demo", bold: true, size: 32 }),
                    ],
                    spacing: { after: 400 },
                }),

                // 1. MathBox
                new Paragraph({
                    children: [new TextRun({ bold: true, text: "1. MathBox", size: 28 })],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathBox({
                                    children: [new MathRun("x + y")],
                                    properties: { opEmu: true },
                                }),
                                new MathRun(" = "),
                                new MathBox({
                                    children: [new MathRun("z")],
                                }),
                            ],
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // 2. MathBorderBox
                new Paragraph({
                    children: [new TextRun({ bold: true, text: "2. MathBorderBox", size: 28 })],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathBorderBox({
                                    children: [new MathRun("a")],
                                }),
                                new MathRun(" + "),
                                new MathBorderBox({
                                    children: [new MathRun("b")],
                                    properties: { hideTop: true, hideBot: true },
                                }),
                            ],
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // 3. Equation Array
                new Paragraph({
                    children: [new TextRun({ bold: true, text: "3. Equation Array", size: 28 })],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathEqArr({
                                    rows: [[new MathRun("x + y = 1")], [new MathRun("2x - y = 3")]],
                                }),
                            ],
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // 4. Group Character
                new Paragraph({
                    children: [
                        new TextRun({ bold: true, text: "4. Group Character (brace)", size: 28 }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathGroupChr({
                                    children: [
                                        new MathEqArr({
                                            rows: [
                                                [new MathRun("a")],
                                                [new MathRun("b")],
                                                [new MathRun("c")],
                                            ],
                                        }),
                                    ],
                                    properties: {
                                        chr: "\u007B",
                                        pos: "bot",
                                        vertJc: "top",
                                    },
                                }),
                            ],
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // 5. Matrix
                new Paragraph({
                    children: [new TextRun({ bold: true, text: "5. Matrix", size: 28 })],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathMatrix({
                                    rows: [
                                        [new MathRun("1"), new MathRun("0")],
                                        [new MathRun("0"), new MathRun("1")],
                                    ],
                                }),
                            ],
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // 6. Phantom
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "6. Phantom (invisible placeholder)",
                            size: 28,
                        }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathRun("f(x) = "),
                                new MathFraction({
                                    numerator: [
                                        new MathPhant({
                                            children: [new MathRun("dy")],
                                            properties: { zeroAsc: true, zeroDesc: true },
                                        }),
                                    ],
                                    denominator: [new MathRun("dx")],
                                }),
                            ],
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // 7. SuperScript
                new Paragraph({
                    children: [
                        new TextRun({ bold: true, text: "7. SuperScript (E = mc²)", size: 28 }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathRun("E = m"),
                                new MathSuperScript({
                                    children: [new MathRun("c")],
                                    superScript: [new MathRun("2")],
                                }),
                            ],
                        }),
                    ],
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
