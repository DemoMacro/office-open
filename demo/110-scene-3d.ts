// Demo: 3D scene (CT_Scene3D) - DrawingML
import * as fs from "fs";

import { Document, Packer, Paragraph, TextRun, WpsShapeRun } from "docx-plus";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [new TextRun({ text: "3D Scene Demo", bold: true, size: 32 })],
                    spacing: { after: 400 },
                }),

                // 1. Shape with 3D scene (camera + light rig)
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "1. Isometric Camera + Three-Point Light",
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
                                        new TextRun({ bold: true, color: "FFFFFF", text: "3D" }),
                                    ],
                                }),
                            ],
                            scene3d: {
                                camera: { preset: "isometricTopUp" },
                                lightRig: { direction: "t", rig: "threePt" },
                            },
                            shape3d: {
                                bevelT: { prst: "CIRCLE", w: 76200, h: 76200 },
                                extrusionColor: { value: "2F5496" },
                                prstMaterial: "PLASTIC",
                                z: 127000,
                            },
                            solidFill: { value: "4472C4" },
                            transformation: { height: 150, width: 300 },
                            type: "wps",
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // 2. Perspective camera
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "2. Perspective Camera + Morning Light",
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
                                            text: "Perspective",
                                        }),
                                    ],
                                }),
                            ],
                            scene3d: {
                                camera: {
                                    fov: 600000,
                                    preset: "perspectiveFront",
                                    zoom: "120000",
                                },
                                lightRig: { direction: "t", rig: "morning" },
                            },
                            shape3d: {
                                bevelB: { prst: "CIRCLE", w: 50800, h: 25400 },
                                bevelT: { prst: "RELAXED_INSET", w: 50800, h: 50800 },
                                extrusionColor: { value: "7030A0" },
                                prstMaterial: "MATTE",
                                z: 76200,
                            },
                            solidFill: { value: "7030A0" },
                            transformation: { height: 150, width: 300 },
                            type: "wps",
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                // 3. Camera with rotation
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "3. Rotated Camera",
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
                                    children: [new TextRun({ bold: true, text: "Rotated View" })],
                                }),
                            ],
                            scene3d: {
                                camera: {
                                    preset: "orthographicFront",
                                    rotation: {
                                        lat: 0,
                                        lon: 2700000,
                                        rev: 450000,
                                    },
                                },
                                lightRig: { direction: "t", rig: "balanced" },
                            },
                            shape3d: {
                                bevelT: { prst: "ANGLE", w: 76200, h: 38100 },
                                contourColor: { value: "000000" },
                                contourW: 12700,
                                extrusionColor: { value: "C00000" },
                                prstMaterial: "WARM_MATTE",
                                z: 50800,
                            },
                            solidFill: { value: "ED7D31" },
                            transformation: { height: 150, width: 300 },
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
