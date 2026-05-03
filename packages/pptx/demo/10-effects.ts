import * as fs from "fs";

import { Presentation, Packer, Shape, SolidFill } from "../src";

const pres = new Presentation({
    title: "Effects Demo",
    creator: "Demo",
    slides: [
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 500,
                    height: 60,
                    text: "Shape Effects Demo",
                    fill: new SolidFill("4472C4"),
                }),
                new Shape({
                    x: 50,
                    y: 120,
                    width: 200,
                    height: 120,
                    text: "Outer Shadow",
                    fill: new SolidFill("ED7D31"),
                    effects: {
                        outerShadow: {
                            blur: 50800,
                            dist: 38100,
                            direction: 5400000,
                            color: "000000",
                            alpha: 50,
                        },
                    },
                }),
                new Shape({
                    x: 300,
                    y: 120,
                    width: 200,
                    height: 120,
                    text: "Glow",
                    fill: new SolidFill("70AD47"),
                    effects: {
                        glow: { radius: 152400, color: "92D050", alpha: 60 },
                    },
                }),
                new Shape({
                    x: 550,
                    y: 120,
                    width: 200,
                    height: 120,
                    text: "Reflection",
                    fill: new SolidFill("FFC000"),
                    effects: {
                        reflection: {
                            blurRadius: 6350,
                            dist: 38100,
                            direction: 5400000,
                            fadeStart: 90,
                            fadeEnd: 0,
                        },
                    },
                }),
                new Shape({
                    x: 50,
                    y: 280,
                    width: 200,
                    height: 120,
                    text: "Inner Shadow",
                    fill: new SolidFill("5B9BD5"),
                    effects: {
                        innerShadow: {
                            blur: 40000,
                            dist: 30000,
                            direction: 5400000,
                            color: "000000",
                            alpha: 40,
                        },
                    },
                }),
                new Shape({
                    x: 300,
                    y: 280,
                    width: 200,
                    height: 120,
                    text: "Soft Edge",
                    fill: new SolidFill("BF8F00"),
                    effects: {
                        softEdge: { radius: 50800 },
                    },
                }),
                new Shape({
                    x: 550,
                    y: 280,
                    width: 200,
                    height: 120,
                    text: "Shadow + Glow",
                    fill: new SolidFill("7030A0"),
                    effects: {
                        outerShadow: {
                            blur: 40000,
                            dist: 30000,
                            direction: 2700000,
                            color: "000000",
                            alpha: 40,
                        },
                        glow: { radius: 101600, color: "B381E7", alpha: 35 },
                    },
                }),
            ],
        },
        // Slide 2: 3D rotation
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 500,
                    height: 50,
                    text: "3D Rotation & Extrusion",
                    fill: new SolidFill("4472C4"),
                }),
                new Shape({
                    x: 50,
                    y: 120,
                    width: 200,
                    height: 200,
                    text: "X=30 Y=0",
                    fill: new SolidFill("4472C4"),
                    effects: { rotation3D: { x: 30 } },
                }),
                new Shape({
                    x: 300,
                    y: 120,
                    width: 200,
                    height: 200,
                    text: "X=0 Y=45",
                    fill: new SolidFill("ED7D31"),
                    effects: { rotation3D: { y: 45 } },
                }),
                new Shape({
                    x: 550,
                    y: 120,
                    width: 200,
                    height: 200,
                    text: "X=20 Y=30 Z=10",
                    fill: new SolidFill("70AD47"),
                    effects: { rotation3D: { x: 20, y: 30, z: 10, perspective: 500 } },
                }),
                new Shape({
                    x: 50,
                    y: 370,
                    width: 200,
                    height: 150,
                    text: "Extruded",
                    fill: new SolidFill("FFC000"),
                    effects: {
                        rotation3D: { x: 25, y: 15 },
                        extrusionH: 50000,
                        material: "plastic",
                    },
                }),
                new Shape({
                    x: 300,
                    y: 370,
                    width: 200,
                    height: 150,
                    text: "Bevel Top",
                    fill: new SolidFill("7030A0"),
                    effects: {
                        rotation3D: { x: 20 },
                        bevelTop: { width: 8, height: 8 },
                        extrusionH: 25000,
                        material: "metal",
                    },
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(pres);
fs.writeFileSync("My Presentation.pptx", buffer);
