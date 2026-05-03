import * as fs from "fs";

import { Background, Presentation, Packer, Shape } from "@office-open/pptx";

const pres = new Presentation({
    title: "Phase 2 Demo",
    creator: "Demo",
    slides: [
        {
            background: new Background({ fill: "F2F2F2" }),
            children: [
                new Shape({
                    x: 50,
                    y: 50,
                    width: 400,
                    height: 100,
                    text: "With Outline",
                    geometry: "roundRect",
                    fill: "FFFFFF",
                    outline: { width: 25400, color: "4472C4", dashStyle: "dash" },
                }),
                new Shape({
                    x: 500,
                    y: 50,
                    width: 400,
                    height: 100,
                    text: "Gradient Fill",
                    fill: {
                        type: "gradient",
                        angle: 0,
                        stops: [
                            { position: 0, color: "4472C4" },
                            { position: 100, color: "ED7D31" },
                        ],
                    },
                }),
            ],
        },
        {
            background: new Background({
                fill: {
                    type: "gradient",
                    angle: 270,
                    stops: [
                        { position: 0, color: "1a1a2e" },
                        { position: 100, color: "16213e" },
                    ],
                },
            }),
            children: [
                new Shape({
                    x: 100,
                    y: 200,
                    width: 300,
                    height: 200,
                    text: "On Gradient BG",
                    fill: "FFFFFF",
                    outline: { width: 12700, color: "FFC000" },
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(pres);
fs.writeFileSync("My Presentation.pptx", buffer);
