// Demo: Custom geometry (CT_CustomGeometry2D) - DrawingML
import * as fs from "fs";

import { Document, Packer, Paragraph, TextRun, WpsShapeRun } from "@office-open/docx";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [new TextRun({ text: "Custom Geometry Demo", bold: true, size: 32 })],
                    spacing: { after: 400 },
                }),

                // 1. Simple triangle
                new Paragraph({
                    children: [new TextRun({ bold: true, text: "1. Triangle", size: 28 })],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new WpsShapeRun({
                            children: [
                                new Paragraph({
                                    alignment: "center",
                                    children: [new TextRun("Triangle")],
                                }),
                            ],
                            customGeometry: {
                                pathList: [
                                    {
                                        w: 100000,
                                        h: 100000,
                                        commands: [
                                            { command: "moveTo", point: { x: "50000", y: "0" } },
                                            {
                                                command: "lineTo",
                                                point: { x: "100000", y: "100000" },
                                            },
                                            { command: "lineTo", point: { x: "0", y: "100000" } },
                                            { command: "close" },
                                        ],
                                    },
                                ],
                            },
                            fill: "4472C4",
                            transformation: { height: 150, width: 200 },
                            type: "wps",
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // 2. Diamond
                new Paragraph({
                    children: [new TextRun({ bold: true, text: "2. Diamond", size: 28 })],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new WpsShapeRun({
                            children: [
                                new Paragraph({
                                    alignment: "center",
                                    children: [new TextRun("Diamond")],
                                }),
                            ],
                            customGeometry: {
                                pathList: [
                                    {
                                        w: 100000,
                                        h: 100000,
                                        commands: [
                                            { command: "moveTo", point: { x: "50000", y: "0" } },
                                            {
                                                command: "lineTo",
                                                point: { x: "100000", y: "50000" },
                                            },
                                            {
                                                command: "lineTo",
                                                point: { x: "50000", y: "100000" },
                                            },
                                            { command: "lineTo", point: { x: "0", y: "50000" } },
                                            { command: "close" },
                                        ],
                                    },
                                ],
                            },
                            outline: {
                                color: { value: "C00000" },
                                type: "solidFill",
                                width: 12700,
                            },
                            fill: "F4B183",
                            transformation: { height: 150, width: 200 },
                            type: "wps",
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // 3. Shape with arc
                new Paragraph({
                    children: [new TextRun({ bold: true, text: "3. Shape with Arc", size: 28 })],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new WpsShapeRun({
                            children: [
                                new Paragraph({
                                    alignment: "center",
                                    children: [new TextRun("Arc Shape")],
                                }),
                            ],
                            customGeometry: {
                                pathList: [
                                    {
                                        w: 100000,
                                        h: 100000,
                                        commands: [
                                            { command: "moveTo", point: { x: "0", y: "50000" } },
                                            {
                                                command: "arcTo",
                                                hR: "50000",
                                                stAng: "0",
                                                swAng: "5400000",
                                                wR: "50000",
                                            },
                                            { command: "lineTo", point: { x: "0", y: "50000" } },
                                            { command: "close" },
                                        ],
                                    },
                                ],
                            },
                            fill: {
                                type: "gradient",
                                options: {
                                    shade: { angle: 5400000 },
                                    stops: [
                                        { color: { value: "002060" }, position: 0 },
                                        { color: { value: "00B0F0" }, position: 100000 },
                                    ],
                                },
                            },
                            outline: {
                                color: { value: "002060" },
                                type: "solidFill",
                                width: 12700,
                            },
                            transformation: { height: 150, width: 200 },
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
