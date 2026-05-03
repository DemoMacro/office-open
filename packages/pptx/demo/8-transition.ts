import * as fs from "fs";

import { Presentation, Packer, Shape, SolidFill } from "@office-open/pptx";

const pres = new Presentation({
    title: "Transition Demo",
    creator: "Demo",
    slides: [
        {
            children: [
                new Shape({
                    x: 100,
                    y: 150,
                    width: 500,
                    height: 100,
                    text: "Fade Transition",
                    fill: new SolidFill("4472C4"),
                }),
            ],
            transition: { type: "fade", speed: "med" },
        },
        {
            children: [
                new Shape({
                    x: 100,
                    y: 150,
                    width: 500,
                    height: 100,
                    text: "Push Transition (Right)",
                    fill: new SolidFill("ED7D31"),
                }),
            ],
            transition: { type: "push", dir: "r", speed: "slow" },
        },
        {
            children: [
                new Shape({
                    x: 100,
                    y: 150,
                    width: 500,
                    height: 100,
                    text: "Wipe Transition (Down)",
                    fill: new SolidFill("70AD47"),
                }),
            ],
            transition: { type: "wipe", dir: "d" },
        },
        {
            children: [
                new Shape({
                    x: 100,
                    y: 150,
                    width: 500,
                    height: 100,
                    text: "Cover Transition (From Right)",
                    fill: new SolidFill("FFC000"),
                }),
            ],
            transition: { type: "cover", dir: "r" },
        },
        {
            children: [
                new Shape({
                    x: 100,
                    y: 150,
                    width: 500,
                    height: 100,
                    text: "Split Transition",
                    fill: new SolidFill("5B9BD5"),
                }),
            ],
            transition: { type: "split", orient: "horz", dir: "out" },
        },
        {
            children: [
                new Shape({
                    x: 100,
                    y: 150,
                    width: 500,
                    height: 100,
                    text: "Wheel Transition (4 spokes)",
                    fill: new SolidFill("BF8F00"),
                }),
            ],
            transition: { type: "wheel", spokes: 4 },
        },
        {
            children: [
                new Shape({
                    x: 100,
                    y: 150,
                    width: 500,
                    height: 100,
                    text: "Dissolve Transition",
                    fill: new SolidFill("7030A0"),
                }),
            ],
            transition: { type: "dissolve", speed: "slow" },
        },
        {
            children: [
                new Shape({
                    x: 100,
                    y: 150,
                    width: 500,
                    height: 100,
                    text: "Random Transition",
                    fill: new SolidFill("C00000"),
                }),
            ],
            transition: { type: "random" },
        },
    ],
});

const buffer = await Packer.toBuffer(pres);
fs.writeFileSync("My Presentation.pptx", buffer);
