import * as fs from "fs";

import { Background, GroupShape, Presentation, Packer, Shape, SolidFill } from "../src";

const pres = new Presentation({
    title: "Group Shape Demo",
    creator: "Demo",
    slides: [
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 400,
                    height: 60,
                    text: "Group Shape Demo",
                    fill: new SolidFill("4472C4"),
                }),
                new GroupShape({
                    x: 50,
                    y: 120,
                    width: 300,
                    height: 200,
                    children: [
                        new Shape({
                            x: 0,
                            y: 0,
                            width: 140,
                            height: 90,
                            text: "Shape A",
                            fill: new SolidFill("ED7D31"),
                        }),
                        new Shape({
                            x: 160,
                            y: 0,
                            width: 140,
                            height: 90,
                            text: "Shape B",
                            fill: new SolidFill("70AD47"),
                        }),
                        new Shape({
                            x: 0,
                            y: 110,
                            width: 300,
                            height: 90,
                            text: "Shape C (wide)",
                            fill: new SolidFill("5B9BD5"),
                        }),
                    ],
                }),
                new GroupShape({
                    x: 400,
                    y: 120,
                    width: 250,
                    height: 200,
                    rotation: 10,
                    children: [
                        new Shape({
                            x: 0,
                            y: 0,
                            width: 250,
                            height: 200,
                            fill: new SolidFill("FFC000"),
                        }),
                        new Shape({
                            x: 25,
                            y: 25,
                            width: 200,
                            height: 150,
                            text: "Rotated Group",
                            fill: new SolidFill("FFFFFF"),
                        }),
                    ],
                }),
            ],
        },
        {
            background: new Background({ fill: new SolidFill("F2F2F2") }),
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 400,
                    height: 60,
                    text: "Nested Groups",
                    fill: new SolidFill("7030A0"),
                }),
                new GroupShape({
                    x: 50,
                    y: 120,
                    width: 600,
                    height: 350,
                    children: [
                        new GroupShape({
                            x: 0,
                            y: 0,
                            width: 280,
                            height: 160,
                            children: [
                                new Shape({
                                    x: 0,
                                    y: 0,
                                    width: 130,
                                    height: 160,
                                    text: "Inner A",
                                    fill: new SolidFill("4472C4"),
                                }),
                                new Shape({
                                    x: 150,
                                    y: 0,
                                    width: 130,
                                    height: 160,
                                    text: "Inner B",
                                    fill: new SolidFill("ED7D31"),
                                }),
                            ],
                        }),
                        new GroupShape({
                            x: 300,
                            y: 0,
                            width: 300,
                            height: 350,
                            children: [
                                new Shape({
                                    x: 0,
                                    y: 0,
                                    width: 300,
                                    height: 160,
                                    text: "Right Top",
                                    fill: new SolidFill("70AD47"),
                                }),
                                new Shape({
                                    x: 0,
                                    y: 180,
                                    width: 140,
                                    height: 170,
                                    text: "RT Bot L",
                                    fill: new SolidFill("FFC000"),
                                }),
                                new Shape({
                                    x: 160,
                                    y: 180,
                                    width: 140,
                                    height: 170,
                                    text: "RT Bot R",
                                    fill: new SolidFill("5B9BD5"),
                                }),
                            ],
                        }),
                    ],
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(pres);
fs.writeFileSync("My Presentation.pptx", buffer);
