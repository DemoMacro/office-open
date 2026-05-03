import * as fs from "fs";

import { Presentation, Packer, Shape } from "@office-open/pptx";

const pres = new Presentation({
    title: "Animation Demo",
    creator: "Demo",
    slides: [
        {
            children: [
                new Shape({
                    x: 50,
                    y: 30,
                    width: 500,
                    height: 60,
                    text: "Shape Animations Demo",
                    fill: "4472C4",
                    animation: { type: "fade", duration: 800 },
                }),
                new Shape({
                    x: 50,
                    y: 120,
                    width: 300,
                    height: 100,
                    text: "Appear",
                    fill: "ED7D31",
                    animation: { type: "appear" },
                }),
                new Shape({
                    x: 400,
                    y: 120,
                    width: 300,
                    height: 100,
                    text: "Fly In (from left)",
                    fill: "70AD47",
                    animation: { type: "fly", direction: "left", duration: 600 },
                }),
                new Shape({
                    x: 50,
                    y: 250,
                    width: 300,
                    height: 100,
                    text: "Wipe (down)",
                    fill: "FFC000",
                    animation: { type: "wipe", direction: "down" },
                }),
                new Shape({
                    x: 400,
                    y: 250,
                    width: 300,
                    height: 100,
                    text: "Dissolve",
                    fill: "5B9BD5",
                    animation: { type: "dissolve", duration: 1000 },
                }),
                new Shape({
                    x: 50,
                    y: 380,
                    width: 300,
                    height: 100,
                    text: "Zoom In",
                    fill: "BF8F00",
                    animation: { type: "zoom" },
                }),
                new Shape({
                    x: 400,
                    y: 380,
                    width: 300,
                    height: 100,
                    text: "Split (horizontal)",
                    fill: "7030A0",
                    animation: { type: "split", direction: "horizontal" },
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(pres);
fs.writeFileSync("My Presentation.pptx", buffer);
