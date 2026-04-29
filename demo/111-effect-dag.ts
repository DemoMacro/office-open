// Demo: Effect DAG container (CT_EffectContainer, a:effectDag) - DrawingML
import * as fs from "fs";

import { Document, Packer, Paragraph, TextRun, WpsShapeRun } from "docx-plus";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [new TextRun({ text: "Effect DAG Demo", bold: true, size: 32 })],
                    spacing: { after: 400 },
                }),

                // 1. Simple effectDag with glow and shadow
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "1. Glow + Outer Shadow",
                            size: 28,
                        }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new WpsShapeRun({
                            children: [
                                new Paragraph({
                                    alignment: "center",
                                    children: [
                                        new TextRun({
                                            bold: true,
                                            color: "FFFFFF",
                                            text: "Glow + Shadow",
                                        }),
                                    ],
                                }),
                            ],
                            effectDag: {
                                glow: { rad: 50800, color: { value: "00B0F0" } },
                                outerShadow: {
                                    blurRad: 76200,
                                    color: { value: "000000" },
                                    dir: 5400000,
                                    dist: 38100,
                                },
                            },
                            solidFill: { value: "0070C0" },
                            transformation: { height: 80, width: 300 },
                            type: "wps",
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // 2. Alpha effects
                new Paragraph({
                    children: [new TextRun({ bold: true, text: "2. Luminance + Tint", size: 28 })],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new WpsShapeRun({
                            children: [
                                new Paragraph({
                                    alignment: "center",
                                    children: [new TextRun("Color Adjusted")],
                                }),
                            ],
                            effectDag: {
                                luminance: { bright: 20000, contrast: -10000 },
                                tint: { hue: 2700000, amt: 30000 },
                            },
                            solidFill: { value: "4472C4" },
                            transformation: { height: 80, width: 300 },
                            type: "wps",
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // 3. Multiple effects with type attribute
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "3. Multiple Effects (sib type)",
                            size: 28,
                        }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new WpsShapeRun({
                            children: [
                                new Paragraph({
                                    alignment: "center",
                                    children: [
                                        new TextRun({
                                            bold: true,
                                            color: "FFFFFF",
                                            text: "Complex Effects",
                                        }),
                                    ],
                                }),
                            ],
                            effectDag: {
                                type: "sib",
                                blur: { rad: 40000 },
                                softEdge: 25400,
                                outerShadow: {
                                    color: { value: "000000" },
                                    blurRad: 50800,
                                    dir: 2700000,
                                    dist: 25400,
                                },
                            },
                            gradientFill: {
                                shade: { angle: 5400000 },
                                stops: [
                                    { color: { value: "1F4E79" }, position: 0 },
                                    { color: { value: "2E75B6" }, position: 100000 },
                                ],
                            },
                            transformation: { height: 80, width: 300 },
                            type: "wps",
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // 4. Effect DAG with container nesting
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "4. Nested Container (tree type)",
                            size: 28,
                        }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new WpsShapeRun({
                            children: [
                                new Paragraph({
                                    alignment: "center",
                                    children: [
                                        new TextRun({
                                            bold: true,
                                            text: "Nested Effects",
                                        }),
                                    ],
                                }),
                            ],
                            effectDag: {
                                type: "tree",
                                containers: [
                                    {
                                        glow: { color: { value: "FF0000" } },
                                        outerShadow: { color: { value: "000000" } },
                                    },
                                ],
                            },
                            outline: {
                                color: { value: "C00000" },
                                type: "solidFill",
                                width: 25400,
                            },
                            solidFill: { value: "F4B183" },
                            transformation: { height: 80, width: 300 },
                            type: "wps",
                        }),
                    ],
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
