import * as fs from "fs";

import { Presentation, Shape, Packer } from "@office-open/pptx";

const pres = new Presentation({
    title: "Theme Demo",
    creator: "Demo",
    masters: [
        {
            theme: {
                name: "Brand Theme",
                colors: {
                    dark1: "1A1A2E",
                    light1: "FFFFFF",
                    dark2: "16213E",
                    light2: "E8E8E8",
                    accent1: "0F3460",
                    accent2: "E94560",
                    accent3: "533483",
                    accent4: "2B2D42",
                    accent5: "8D99AE",
                    accent6: "EDF2F4",
                },
                fonts: {
                    majorFont: "Segoe UI",
                    minorFont: "Segoe UI",
                },
            },
        },
    ],
    slides: [
        {
            children: [
                new Shape({
                    x: 100,
                    y: 100,
                    width: 600,
                    height: 300,
                    text: "Custom Theme",
                    geometry: "rect",
                    fill: "0F3460",
                }),
                new Shape({
                    x: 100,
                    y: 450,
                    width: 600,
                    height: 100,
                    text: "Accent2 highlight",
                    geometry: "rect",
                    fill: "E94560",
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(pres);
fs.writeFileSync("My Presentation.pptx", buffer);
