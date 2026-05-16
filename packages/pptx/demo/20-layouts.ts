import * as fs from "fs";

import { Presentation, Shape, Packer } from "@office-open/pptx";

const pres = new Presentation({
    title: "Layouts Demo",
    creator: "Demo",
    masters: [
        {
            background: {
                fill: "1B2A4A",
            },
            children: [
                new Shape({
                    x: 0,
                    y: 695,
                    width: 1280,
                    height: 25,
                    geometry: "rect",
                    fill: "4472C4",
                }),
            ],
            placeholders: {
                slideNumber: {
                    x: 904,
                    y: 667,
                    width: 288,
                    height: 38,
                },
            },
        },
    ],
    slides: [
        {
            layout: "title",
            children: [
                new Shape({
                    x: 100,
                    y: 100,
                    width: 600,
                    height: 200,
                    text: "Title Slide Layout",
                    geometry: "rect",
                    fill: "0F3460",
                }),
            ],
        },
        {
            layout: "obj",
            children: [
                new Shape({
                    x: 100,
                    y: 100,
                    width: 600,
                    height: 100,
                    text: "Title and Content Layout",
                    fill: "4472C4",
                }),
                new Shape({
                    x: 100,
                    y: 250,
                    width: 600,
                    height: 200,
                    text: "Content area",
                    fill: "E8E8E8",
                }),
            ],
        },
        {
            layout: "twoColTx",
            children: [
                new Shape({
                    x: 100,
                    y: 100,
                    width: 600,
                    height: 100,
                    text: "Two Column Layout",
                    fill: "70AD47",
                }),
            ],
        },
        {
            layout: "blank",
            children: [
                new Shape({
                    x: 100,
                    y: 100,
                    width: 600,
                    height: 300,
                    text: "Blank Layout",
                    fill: "FFC000",
                }),
            ],
        },
        {
            layout: "titleOnly",
            children: [
                new Shape({
                    x: 100,
                    y: 100,
                    width: 600,
                    height: 100,
                    text: "Title Only Layout",
                    fill: "ED7D31",
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(pres);
fs.writeFileSync("My Presentation.pptx", buffer);
