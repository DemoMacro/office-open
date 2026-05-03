import * as fs from "fs";

import { Presentation, Packer, Shape } from "@office-open/pptx";

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
                    fill: "4472C4",
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
                    fill: "ED7D31",
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
                    fill: "70AD47",
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
                    fill: "FFC000",
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
                    fill: "5B9BD5",
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
                    fill: "BF8F00",
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
                    fill: "7030A0",
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
                    fill: "C00000",
                }),
            ],
            transition: { type: "random" },
        },
    ],
});

const buffer = await Packer.toBuffer(pres);
fs.writeFileSync("My Presentation.pptx", buffer);
